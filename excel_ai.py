"""excel_ai: converte uma planilha Excel em texto compacto para uma IA ler.

Uso:   python excel_ai.py planilha.xlsx     (gera planilha.ai.txt ao lado)
Teste: python excel_ai.py                   (sem argumento roda o auto-teste)

Roda 100% local. Unica dependencia: openpyxl (pip install openpyxl).
"""
import datetime
import io
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

LIMITE_NUMEROS = 10    # sequencias de numeros/datas maiores que isso viram resumo
LIMITE_TEXTOS = 50     # idem para textos (menor que isso costuma ser rotulo; fica inteiro)
AMOSTRA = 3            # quantos primeiros valores o resumo mostra
RATIO_DUPLICADA = 0.8  # fracao de formulas iguais para tratar a aba como copia
ERROS = ("#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")

CELL = r"\$?[A-Z]{1,3}\$?\d+"
REF_OK = re.compile(r"^(?:%s(?::%s)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)$" % (CELL, CELL))
PECA_CELULA = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")
PECA_COLUNA = re.compile(r"^(\$?)([A-Z]{1,3})$")
PECA_LINHA = re.compile(r"^(\$?)(\d+)$")
# ponytail: regex pode achar nome de aba dentro de texto entre aspas; raro, so afeta o mapa.
NOME_ABA = r"'((?:[^']|'')+)'|([^\W\d][\w.]*)"  # 'Aba com espaco' ou Aba
RE_ABA = re.compile(r"(?<![#\w.\]])(?:%s)!" % NOME_ABA)
RE_BLOCO = re.compile(r"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))? (=.*)$")
RE_MATRIZ = re.compile(r"^([A-Z]+\d+) array \S+ \{(=.*)\}$")
RE_EXCECAO = re.compile(r"^  except ([A-Z]+\d+) (.*)$")
RE_SUMIU = re.compile(r"^([A-Z]+\d+) no formula$")
RE_REF_QUEBRADA = re.compile(r"(?:(?:%s)!)?#REF!" % NOME_ABA)


def tokens(texto):
    # ponytail: ~4 caracteres por token (estimativa). Use tiktoken se precisar de precisao.
    return len(texto) // 4


def ref(r, c):
    return "%s%d" % (get_column_letter(c), r)


def limpa(formula):
    """Tira prefixos internos do Excel: _xlfn.XLOOKUP vira XLOOKUP."""
    return re.sub(r"_xlfn\.|_xlws\.|_xlpm\.", "", formula)


def fmt(v):
    if isinstance(v, bool):
        return str(v).upper()
    if isinstance(v, float):
        return "%.10g" % v
    if isinstance(v, datetime.datetime):
        return v.date().isoformat() if v.time() == datetime.time(0) else v.isoformat()
    if isinstance(v, (datetime.date, datetime.time)):
        return v.isoformat()
    return str(v).replace("\r", " ").replace("\n", " ")


def formula_de(v):
    """('f', texto) formula comum, ('a', texto, faixa) matriz, ('d',) tabela de dados, None constante."""
    nome = type(v).__name__  # por nome: DataTableFormula nao existe em openpyxl < 3.1
    if nome == "ArrayFormula":
        texto = v.text if v.text.startswith("=") else "=" + v.text
        return ("a", texto, v.ref)
    if nome == "DataTableFormula":
        return ("d",)
    if isinstance(v, str) and v.startswith("=") and len(v) > 1:
        return ("f", v)
    return None


def peca(p, row, col):
    m = PECA_CELULA.match(p)
    if m:
        cabs, letra, rabs, num = m.groups()
        c, r = column_index_from_string(letra), int(num)
        return ("R%d" % r if rabs else "R[%d]" % (r - row)) + ("C%d" % c if cabs else "C[%d]" % (c - col))
    m = PECA_COLUNA.match(p)
    if m:
        c = column_index_from_string(m.group(2))
        return "C%d" % c if m.group(1) else "C[%d]" % (c - col)
    m = PECA_LINHA.match(p)
    r = int(m.group(2))
    return "R%d" % r if m.group(1) else "R[%d]" % (r - row)


def relativa(formula, row, col):
    """=B2*C2 em E2 vira =R[0]C[-3]*R[0]C[-2]: mesmo texto para a coluna inteira."""
    try:
        itens = Tokenizer(formula).items
    except Exception:
        return formula
    partes = []
    for t in itens:
        v = t.value
        # "[" = referencia estruturada (Tabela1[Col]) ou outro arquivo: fica como esta.
        if t.type == Token.OPERAND and t.subtype == Token.RANGE and "[" not in v:
            aba, sep, alvo = v.rpartition("!")
            if REF_OK.match(alvo):
                v = aba + sep + ":".join(peca(p, row, col) for p in alvo.split(":"))
        partes.append(v)
    return "=" + "".join(partes)


def blocos(formulas, valores, vertical, buraco="blank"):
    """Junta celulas vizinhas com o mesmo padrao. Uma celula diferente no meio vira 'except'."""
    por_linha = defaultdict(list)
    for (r, c) in formulas:
        por_linha[c if vertical else r].append(r if vertical else c)

    saida = []
    for fixo in sorted(por_linha):
        def cel(p):
            return (p, fixo) if vertical else (fixo, p)

        segs = []  # [tipo, chave, inicio, fim, excecoes]

        def junta(tipo, chave, p0, p1):
            if segs and tipo in "fx" and segs[-1][0] == tipo and segs[-1][1] == chave and segs[-1][3] == p0 - 1:
                segs[-1][3] = p1
            else:
                segs.append([tipo, chave, p0, p1, []])

        ant = None
        for p in sorted(por_linha[fixo]):
            if ant is not None and p > ant + 1:
                junta("x", None, ant + 1, p - 1)  # buraco: vazio ou constante
            f = formulas[cel(p)]
            if f[0] == "f":
                junta("f", f[2], p, p)
            else:
                junta(f[0], p, p, p)  # matriz e tabela de dados nunca se juntam
            ant = p

        pilha = []
        for s in segs:
            if s[0] == "f" and len(pilha) >= 2:
                meio, antes = pilha[-1], pilha[-2]
                if antes[0] == "f" and antes[1] == s[1] and meio[2] == meio[3] and meio[0] in "fx":
                    pilha.pop()
                    antes[4].append(meio[2])
                    antes[4].extend(s[4])
                    antes[3] = s[3]
                    continue
            pilha.append(s)

        for tipo, _, p0, p1, excecoes in pilha:
            ini = cel(p0)
            if tipo == "a":
                saida.append(linha_formula(ini, formulas[ini]))
            if tipo != "f":
                continue
            faixa = ref(*ini) if p0 == p1 else ref(*ini) + ":" + ref(*cel(p1))
            saida.append("%s %s" % (faixa, limpa(formulas[ini][1])))
            for p in excecoes:
                f = formulas.get(cel(p))
                if f and f[0] == "f":
                    saida.append("  except %s %s" % (ref(*cel(p)), limpa(f[1])))
                elif cel(p) in valores:
                    saida.append("  except %s value %s" % (ref(*cel(p)), fmt(valores[cel(p)])))
                else:
                    saida.append("  except %s %s" % (ref(*cel(p)), buraco))
    return saida


def expandir(linhas, base=None):
    """Le as linhas de formulas de volta para {celula: formula}. Prova que nada se perdeu."""
    out = {}
    if linhas and linhas[0].startswith("same formulas as"):
        out = dict(base)
        resto = []
        for linha in linhas[1:]:
            m = RE_SUMIU.match(linha)
            if m:
                out.pop(m.group(1), None)
            else:
                resto.append(linha)
        out.update(expandir(resto))  # 'except ... unchanged' nao mexe na celula da base
        return out
    for linha in linhas:
        m = RE_EXCECAO.match(linha)
        if m:
            if m.group(2).startswith("="):
                out[m.group(1)] = m.group(2)
            else:
                out.pop(m.group(1), None)
            continue
        m = RE_MATRIZ.match(linha)
        if m:
            out[m.group(1)] = m.group(2)
            continue
        l1, n1, l2, n2, f = RE_BLOCO.match(linha).groups()
        origem = l1 + n1
        c1, r1 = column_index_from_string(l1), int(n1)
        c2, r2 = (column_index_from_string(l2), int(n2)) if l2 else (c1, r1)
        tr = tradutor(f, origem)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                destino = ref(r, c)
                out[destino] = f if destino == origem else tr(destino)
    return out


def tradutor(f, origem):
    """Translator do openpyxl, protegendo ABA!#REF! (o Translator apaga esse pedaco)."""
    guardados = []

    def guarda(m):
        guardados.append(m.group(0))
        return "XLREFERR%dX" % (len(guardados) - 1)

    tr = Translator(RE_REF_QUEBRADA.sub(guarda, f), origin=origem)
    return lambda destino: re.sub(r"XLREFERR(\d+)X", lambda m: guardados[int(m.group(1))],
                                  tr.translate_formula(destino))


def esperado(formulas):
    return {ref(r, c): limpa(f[1]) for (r, c), f in formulas.items() if f[0] in "fa"}


def secao_formulas(formulas, valores, buraco="blank"):
    if not formulas:
        return []
    return min(blocos(formulas, valores, True, buraco), blocos(formulas, valores, False, buraco),
               key=lambda b: len("\n".join(b)))


def por_celula(formulas):
    """Plano B: uma linha por formula (sempre sem perda)."""
    return [linha_formula(pos, formulas[pos]) for pos in sorted(formulas) if formulas[pos][0] in "fa"]


def linha_formula(pos, f):
    """Uma celula na secao de formulas, no formato que expandir() le."""
    if f[0] == "a":
        return "%s array %s {%s}" % (ref(*pos), f[2], limpa(f[1]))
    return "%s %s" % (ref(*pos), limpa(f[1]))


def diferenca(formulas, meu, base_nome, base):
    """'same formulas as X' + so as celulas que mudam, tambem agrupadas em blocos."""
    mudou = {pos: f for pos, f in formulas.items()
             if f[0] in "fa" and base.get(ref(*pos)) != limpa(f[1])}
    sumiu = sorted((k for k in base if k not in meu), key=lambda k: (coordinate_from_string(k)[1], k))
    return (["same formulas as '%s', except:" % base_nome]
            + secao_formulas(mudou, {}, buraco="unchanged")
            + ["%s no formula" % k for k in sumiu])


def tipo_valor(v):
    if isinstance(v, bool):
        return "texts"
    if isinstance(v, (int, float)):
        return "numbers"
    if isinstance(v, (datetime.date, datetime.time)):
        return "dates"
    return "texts"


def secao_dados(valores):
    """Rotulos ficam inteiros. Sequencias longas de numeros/datas/textos viram resumo."""
    def itens(vertical):
        out = []
        por_linha = defaultdict(list)
        for (r, c) in valores:
            por_linha[c if vertical else r].append(r if vertical else c)
        for fixo, pos in por_linha.items():
            pos.sort()
            cels_de = [(p, fixo) if vertical else (fixo, p) for p in pos]
            corridas = [[cels_de[0]]]
            for ant, x in zip(cels_de, cels_de[1:]):
                vizinha = (x[0] - ant[0] + x[1] - ant[1]) == 1
                if vizinha and tipo_valor(valores[x]) == tipo_valor(valores[ant]):
                    corridas[-1].append(x)
                else:
                    corridas.append([x])
            for cels in corridas:
                vals = [valores[x] for x in cels]
                tipo = tipo_valor(vals[0])
                faixa = "%s:%s" % (ref(*cels[0]), ref(*cels[-1]))
                primeiros = ", ".join(fmt(v) for v in vals[:AMOSTRA])
                if tipo != "texts" and len(cels) > LIMITE_NUMEROS:
                    try:
                        menor, maior = fmt(min(vals)), fmt(max(vals))
                    except TypeError:  # data misturada com hora
                        menor = maior = "?"
                    out.append(cels[0] + ("%s [%d %s] min %s max %s first: %s" % (
                        faixa, len(vals), tipo, menor, maior, primeiros), True))
                elif tipo == "texts" and len(cels) > LIMITE_TEXTOS:
                    out.append(cels[0] + ("%s [%d texts, %d distinct] first: %s" % (
                        faixa, len(vals), len(set(map(fmt, vals))), primeiros), True))
                else:
                    i = 0
                    while i < len(cels):  # textos iguais seguidos viram uma linha so
                        j = i
                        while j + 1 < len(cels) and fmt(vals[j + 1]) == fmt(vals[i]):
                            j += 1
                        if j - i >= 2:
                            out.append(cels[i] + ("%s:%s %s (x%d)" % (
                                ref(*cels[i]), ref(*cels[j]), fmt(vals[i]), j - i + 1), True))
                        else:
                            out.extend(x + ("%s %s" % (ref(*x), fmt(valores[x])), False) for x in cels[i:j + 1])
                        i = j + 1
        return out

    escolhido = min(itens(True), itens(False), key=len)
    escolhido.sort()
    linhas, linha_atual, r_atual = [], [], None
    for r, c, texto, resumo in escolhido:
        if resumo or r != r_atual:
            if linha_atual:
                linhas.append(" | ".join(linha_atual))
            linha_atual, r_atual = [], r
        if resumo:
            linhas.append(texto)
            r_atual = None
        else:
            linha_atual.append(texto)
    if linha_atual:
        linhas.append(" | ".join(linha_atual))
    return linhas


def tem_vba(fonte):
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    try:
        with zipfile.ZipFile(fonte) as z:
            return "xl/vbaProject.bin" in z.namelist()
    except zipfile.BadZipFile:
        return False


def nomes_definidos(wb):
    dn = wb.defined_names
    lista = list(dn.values()) if hasattr(dn, "values") else list(getattr(dn, "definedName", []))
    for ws in wb.worksheets:  # nomes com escopo de aba (openpyxl 3.1+)
        lista.extend(getattr(ws, "defined_names", {}).values())
    return ["%s = %s" % (d.name, d.attr_text) for d in lista]


def converter(fonte):
    """Retorna (texto, numeros). fonte = caminho ou arquivo aberto."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    wb = load_workbook(fonte)
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    wb_v = load_workbook(fonte, read_only=True, data_only=True)

    avisos, mapa, corpo = [], [], []
    bruto_chars = 0  # tamanho do dump celula a celula (so para a comparacao)
    falhas, total_formulas, sem_cache = [], 0, True
    ja_vistas = []  # (nome, esperado) das abas anteriores
    s_form = s_dados = 0

    for ws in wb.worksheets:
        cache = {}
        for linha in wb_v[ws.title].iter_rows():
            for c in linha:
                if c.value is not None:
                    cache[(c.row, c.column)] = c.value
        formulas, valores = {}, {}
        for linha in ws.iter_rows():
            for c in linha:
                if c.value is None:
                    continue
                f = formula_de(c.value)
                pos = (c.row, c.column)
                if f is None:
                    valores[pos] = c.value
                    bruto_chars += len("%s\t%s" % (ref(*pos), fmt(c.value))) + 1
                else:
                    if f[0] == "f":
                        f = f + (relativa(f[1], c.row, c.column),)
                    formulas[pos] = f
                    bruto_chars += len("%s\t%s\t%s" % (ref(*pos), f[1] if len(f) > 1 else "TABLE()",
                                                      cache.get(pos))) + 1
        if any(pos in cache for pos in formulas):
            sem_cache = False

        textos = [f[1] for f in formulas.values() if f[0] in "fa"]
        total_formulas += len(textos)
        meu = esperado(formulas)

        # avisos da aba
        n_ref = sum(1 for t in textos if "#REF!" in t)
        n_erro = sum(1 for p in formulas if isinstance(cache.get(p), str) and cache[p] in ERROS)
        n_dinamica = sum(1 for t in textos if re.search(r"\b(OFFSET|INDIRECT)\(", t))
        n_tabela_dados = sum(1 for f in formulas.values() if f[0] == "d")
        nome = "'%s'" % ws.title
        if n_ref:
            avisos.append("- %s: %d formulas contain #REF! (broken reference)" % (nome, n_ref))
        if n_erro:
            avisos.append("- %s: %d formulas show an error value (#N/A, #DIV/0!...)" % (nome, n_erro))
        if n_dinamica:
            avisos.append("- %s: %d formulas use OFFSET/INDIRECT (reference decided at run time)" % (nome, n_dinamica))
        if n_tabela_dados:
            avisos.append("- %s: %d data-table (what-if) cells, not shown as formulas" % (nome, n_tabela_dados))

        # mapa
        usa, externo = Counter(), 0
        for t in textos:
            achadas = set()
            for q, s in RE_ABA.findall(t):
                achadas.add(q.replace("''", "'") if q else s)
            for a in achadas:
                if a != ws.title:
                    usa[a] += 1
            externo += bool(re.search(r"\[\d+\]", t))
        partes = ["- %s %s" % (nome, ws.dimensions)]
        if ws.sheet_state != "visible":
            partes.append("HIDDEN")
        partes.append("%d formulas" % len(textos))
        if usa:
            partes.append("uses " + ", ".join("'%s' (%d)" % kv for kv in usa.most_common()))
        if externo:
            partes.append("%d formulas point to other files" % externo)
        mapa.append(", ".join(partes))
        for t in ws.tables.values():
            colunas = [col.name for col in t.tableColumns]
            mapa.append("  table %s %s columns: %s" % (t.displayName, t.ref, ", ".join(colunas)))

        # formulas
        linhas_f = secao_formulas(formulas, valores)
        if expandir(linhas_f) != meu:
            falhas.append(ws.title)
            linhas_f = por_celula(formulas)
        for base_nome, base in ja_vistas:
            iguais = sum(1 for k, v in meu.items() if base.get(k) == v)
            uniao = len(set(meu) | set(base))
            if uniao and iguais / float(uniao) >= RATIO_DUPLICADA:
                dif = diferenca(formulas, meu, base_nome, base)
                if len("\n".join(dif)) < len("\n".join(linhas_f)) and expandir(dif, base) == meu:
                    linhas_f = dif
                    break
        ja_vistas.append((ws.title, meu))

        linhas_d = secao_dados(valores)
        bloco = ["", "# Sheet %s" % nome]
        if linhas_f:
            bloco += ["## Formulas"] + linhas_f
        if linhas_d:
            bloco += ["## Data"] + linhas_d
        s_form += tokens("\n".join(linhas_f))
        s_dados += tokens("\n".join(linhas_d))
        corpo += bloco

    wb_v.close()  # read_only segura o arquivo aberto
    if tem_vba(fonte):
        avisos.insert(0, "- File has VBA macros: some values may be written by a macro, not by formulas.")
    if total_formulas and sem_cache:
        avisos.insert(0, "- No calculated values saved in the file (never opened in Excel?): error checks are blind.")
    if falhas:
        avisos.insert(0, "- WARNING: compact view failed the self-check on %d sheet(s): %s. "
                         "Those sheets list one formula per cell." % (len(falhas), ", ".join(falhas)))
    nomes = nomes_definidos(wb)
    if nomes:
        mapa += ["Defined names:"] + ["  " + n for n in nomes]

    cabeca = [
        "# excel_ai dump",
        "How to read formulas: '<range> <formula>' gives the formula of the first cell;",
        "the rest of the range is the same formula copied (Excel fill), refs shifting as usual.",
        "'except <cell> ...' overrides one cell inside the range above it.",
        "'same formulas as X' means the sheet is a copy of X except the listed cells.",
    ]
    texto_avisos = ["", "## Warnings"] + (avisos or ["- none"])
    texto_mapa = ["", "## Map"] + mapa
    texto = "\n".join(cabeca + texto_avisos + texto_mapa + corpo) + "\n"
    numeros = {
        "bruto": max(bruto_chars - 1, 0) // 4,  # = tokens() do dump juntado com "\n"
        "novo": tokens(texto),
        "avisos": tokens("\n".join(texto_avisos)),
        "mapa": tokens("\n".join(texto_mapa)),
        "formulas": s_form,
        "dados": s_dados,
        "total_formulas": total_formulas,
        "falhas": len(falhas),
    }
    return texto, numeros


def main(caminho):
    if not os.path.isfile(caminho):
        sys.stderr.write("erro: arquivo nao encontrado: %s\n" % ascii(caminho))
        return 2
    try:
        texto, n = converter(caminho)
    except Exception as e:
        sys.stderr.write("erro: nao consegui ler a planilha (%s: %s)\n" % (type(e).__name__, ascii(str(e))))
        return 2
    destino = os.path.splitext(caminho)[0] + ".ai.txt"
    with io.open(destino, "w", encoding="utf-8", newline="\n") as arq:
        arq.write(texto)
    print("Tokens hoje (celula a celula): %10s" % "{:,}".format(n["bruto"]))
    print("Tokens com excel_ai:           %10s" % "{:,}".format(n["novo"]))
    print("  avisos %s | mapa %s | formulas %s | dados %s" % tuple(
        "{:,}".format(n[k]) for k in ("avisos", "mapa", "formulas", "dados")))
    if n["novo"]:
        print("Reducao: %.1fx" % (n["bruto"] / float(n["novo"])))
    if n["falhas"]:
        print("Conferencia das formulas: FALHOU em %d aba(s); essas abas sairam celula a celula." % n["falhas"])
    else:
        print("Conferencia das formulas: 100%% (%s formulas sem perda)" % "{:,}".format(n["total_formulas"]))
    print("Arquivo gerado: %s" % os.path.basename(destino).encode("ascii", "backslashreplace").decode())
    return 1 if n["falhas"] else 0


def _fixture():
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.formula import ArrayFormula
    from openpyxl.worksheet.table import Table

    wb = Workbook()
    p = wb.active
    p.title = "Premissas"
    p["A1"], p["B1"] = "Taxa", 0.012
    p["A2"], p["B2"] = "Principal", 1000000
    p["A3"], p["B3"] = "Carencia", 6

    f = wb.create_sheet("Fluxo")
    f.append(["Parcela", "Saldo", "Taxa", "Juros"])
    for r in range(2, 102):
        f["A%d" % r] = r - 1
        f["B%d" % r] = "=Premissas!$B$2-A%d" % r
        f["C%d" % r] = "=_xlfn.XLOOKUP(\"Taxa\",Premissas!$A:$A,Premissas!$B:$B)"
        f["D%d" % r] = "=B%d*C%d" % (r, r)
    f["D50"] = "=B50*0"   # formula diferente no meio
    f["D60"] = 7          # numero digitado no meio
    f["D70"] = None       # vazio no meio
    f["D80"], f["D81"] = "=1", "=2"  # duas diferentes juntas
    f["F2"] = ArrayFormula("F2:F4", "=B2:B4*2")
    f["G2"] = "=OFFSET(A1,1,0)"
    f["H2"] = "=#REF!*2"
    f["I2"] = "=SUM(Tabela1[Valor])"
    f["J2"] = "='Plan Extra 2'!A1+TAXA2024"

    m = wb.create_sheet("Meses")
    for c in range(2, 14):
        m.cell(row=1, column=c, value=c * 10)
        m.cell(row=2, column=c, value="=%s1*2" % get_column_letter(c))

    t = wb.create_sheet("Plan Extra 2")
    t.append(["Nome", "Valor"])
    for i in range(3):
        t.append(["item%d" % i, i])
    t.add_table(Table(displayName="Tabela1", ref="A1:B4"))
    t.sheet_state = "hidden"

    copia = wb.copy_worksheet(f)
    copia.title = "Fluxo B"
    copia["D30"] = "=B30*C30*2"
    copia["A2"] = 999

    nome = DefinedName("TAXA2024", attr_text="Premissas!$B$1")
    try:
        wb.defined_names.add(nome)
    except AttributeError:
        wb.defined_names.append(nome)
    buf = io.BytesIO()
    wb.save(buf)
    return buf


def autoteste():
    # chave relativa: mesmo padrao na coluna inteira
    assert relativa("=B2*C2", 2, 5) == "=R[0]C[-3]*R[0]C[-2]"
    assert relativa("=B3*C3", 3, 5) == relativa("=B2*C2", 2, 5)
    assert relativa("=B2*$A$1", 2, 3) == "=R[0]C[-1]*R1C1"
    assert relativa("=$B2*B$1", 2, 3) == "=R[0]C2*R1C[-1]"
    assert relativa("='Plan 1'!B2", 2, 2) == "='Plan 1'!R[0]C[0]"
    assert relativa('=IF(A2="B2",1,0)', 2, 2) == '=IF(R[0]C[-1]="B2",1,0)'
    assert relativa("=Tabela1[CDI2024]", 2, 2) == "=Tabela1[CDI2024]"
    assert relativa("=Tabela1[[#This Row],[A1]]", 2, 2) == "=Tabela1[[#This Row],[A1]]"
    assert relativa("=TAXA2024*A2", 2, 2) == "=TAXA2024*R[0]C[-1]"
    assert relativa("=SUM(B:B)", 1, 3) == relativa("=SUM(C:C)", 1, 4)
    assert relativa("=SUM(2:2)", 2, 1) == relativa("=SUM(3:3)", 3, 1)
    assert relativa("=SUM($B:$B)", 1, 3) != relativa("=SUM($C:$C)", 1, 4)
    assert limpa("=_xlfn.XLOOKUP(A1,B:B,C:C)") == "=XLOOKUP(A1,B:B,C:C)"

    texto, n = converter(_fixture())
    secoes = texto.split("\n# Sheet ")
    fluxo = [s for s in secoes if s.startswith("'Fluxo'")][0]
    # bloco com excecoes: formula diferente, numero e vazio no meio
    assert "D2:D79 =B2*C2" in fluxo, fluxo
    assert "  except D50 =B50*0" in fluxo
    assert "  except D60 value 7" in fluxo
    assert "  except D70 blank" in fluxo
    assert "D80 =1" in fluxo and "D81 =2" in fluxo and "D82:D101 =B82*C82" in fluxo
    assert "C2:C101 =XLOOKUP(" in fluxo and "_xlfn" not in fluxo
    assert "F2 array F2:F4 {=B2:B4*2}" in fluxo
    assert "I2 =SUM(Tabela1[Valor])" in fluxo
    # meses nas colunas: agrupa na horizontal
    assert "B2:M2 =B1*2" in texto
    # copia de aba: so a diferenca, mas os dados ficam inteiros
    copia = [s for s in secoes if s.startswith("'Fluxo B'")][0]
    assert "same formulas as 'Fluxo', except:" in copia and "D30 =B30*C30*2" in copia
    # diferencas tambem agrupadas; celula igual a base no meio vira 'unchanged'
    dif = diferenca({(r, 1): ("f", "=X%d" % r, relativa("=X%d" % r, r, 1)) for r in range(2, 9)},
                    {"A%d" % r: "=X%d" % r for r in range(2, 9)}, "B",
                    {"A5": "=X5", "A9": "=Y"})
    assert dif == ["same formulas as 'B', except:", "A2:A8 =X2", "  except A5 unchanged", "A9 no formula"], dif
    assert expandir(dif, {"A5": "=X5", "A9": "=Y"}) == {"A%d" % r: "=X%d" % r for r in range(2, 9)}
    # textos iguais seguidos viram uma linha; coluna longa de textos vira resumo
    linhas = secao_dados({(r, 1): "INSS" for r in range(2, 8)})
    assert linhas == ["A2:A7 INSS (x6)"], linhas
    linhas = secao_dados({(r, 1): "cod%d" % r for r in range(2, 70)})
    assert linhas == ["A2:A69 [68 texts, 68 distinct] first: cod2, cod3, cod4"], linhas
    assert "A2:A101 [100 numbers] min 2 max 999 first: 999, 2, 3" in copia
    # dados: texto inteiro, sequencia longa de numeros resumida, premissas inteiras
    assert "A1 Parcela | B1 Saldo | C1 Taxa | D1 Juros" in fluxo
    assert "A2:A101 [100 numbers] min 1 max 100 first: 1, 2, 3" in fluxo
    assert "A1 Taxa | B1 0.012" in texto
    # mapa e avisos
    assert "- 'Fluxo' A1:J101, 303 formulas, uses 'Premissas' (200), 'Plan Extra 2' (1)" in texto
    assert "'Plan Extra 2' A1:B4, HIDDEN" in texto
    assert "  table Tabela1 A1:B4 columns: Nome, Valor" in texto
    assert "  TAXA2024 = Premissas!$B$1" in texto
    assert "- 'Fluxo': 1 formulas contain #REF!" in texto
    assert "- 'Fluxo': 1 formulas use OFFSET/INDIRECT" in texto
    assert "No calculated values saved" in texto
    assert "VBA" not in texto
    assert n["falhas"] == 0 and n["total_formulas"] > 0

    # conferencia pega uma linha estragada
    linhas = ["D2:D10 =B2*C2"]
    certo = {"D%d" % r: "=B%d*C%d" % (r, r) for r in range(2, 11)}
    assert expandir(linhas) == certo
    assert expandir(["D2:D10 =B2*C3"]) != certo
    # Translator do openpyxl apagava ABA!#REF!; tradutor() protege
    assert tradutor("=XLOOKUP(A2,MIX!#REF!,'P 1'!#REF!)+#REF!", "B2")("B3") == "=XLOOKUP(A3,MIX!#REF!,'P 1'!#REF!)+#REF!"

    # macro: zip com vbaProject.bin
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as arq:
        arq.writestr("xl/vbaProject.bin", b"x")
    assert tem_vba(z) and not tem_vba(_fixture())
    print("auto-teste ok")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(main(sys.argv[1]))
    autoteste()
