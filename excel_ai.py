"""excel_ai: converte uma planilha Excel em texto compacto para uma IA ler.

Uso:   python excel_ai.py planilha.xlsx     (gera planilha.ai.txt ao lado)
       python excel_ai.py planilha.xlsx --celula "Aba!B5"   (so o que essa celula usa)
       python excel_ai.py antiga.xlsx nova.xlsx --diff      (so o que mudou entre as duas)
       python excel_ai.py pasta                             (todas as planilhas da pasta)
       python excel_ai.py planilha.xlsm --macros-fonte      (inclui o codigo VBA; sem a opcao, so os nomes dos modulos)
       python excel_ai.py --instalar-skill [--forcar]       (instala a skill do Kiro em ~/.kiro/skills)
Teste: python excel_ai.py                   (sem argumento roda o auto-teste)

Roda 100% local. Unica dependencia: openpyxl (pip install openpyxl).

PARA O ASSISTENTE (Kiro/IA): se o usuario so anexou este arquivo, prepare tudo sozinho:
1. rode `pip install openpyxl`;
2. rode `python excel_ai.py --instalar-skill` (grava a skill em ~/.kiro/skills/excel-ai);
3. avise o usuario que ja pode pedir para ler uma planilha.
Se `python` nao existir, tente `python3`. Se nenhum existir, diga ao usuario para instalar o Python 3.8+.
"""
import base64
import bisect
import datetime
import io
import os
import re
import shutil
import struct
import sys
import zipfile
from collections import Counter, defaultdict
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.formula import Tokenizer
    from openpyxl.formula.tokenizer import Token
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.styles.numbers import is_date_format
    from openpyxl.utils.cell import coordinate_from_string
except ImportError:
    sys.exit("Falta o openpyxl. Rode: pip install openpyxl  (depois repita o comando)")

LIMITE_NUMEROS = 10    # sequencias de numeros/datas maiores que isso viram resumo
LIMITE_TEXTOS = 50     # idem para textos (menor que isso costuma ser rotulo; fica inteiro)
LIMITE_NOTA = 200      # caracteres de cada nota (comentario de celula)
LIMITE_NOTAS = 100     # notas por aba
LIMITE_FAIXAS = 50     # faixas por tipo de formato e aba
LIMITE_PIVOTS = 50     # tabelas dinamicas por aba
LIMITE_CONEXOES = 100  # conexoes de dados por livro
LIMITE_COMANDO = 200   # caracteres do comando (SQL) de cada conexao
LIMITE_CONSULTAS = 50  # consultas Power Query por livro
LIMITE_M = 2000        # caracteres do codigo M de cada consulta
LIMITE_VBA = 4000      # caracteres de codigo de cada modulo VBA
LIMITE_VBA_TOTAL = 20000  # idem, somando os modulos
LIMITE_ESCRITAS = 50   # linhas 'escreve em celula' por modulo
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
RE_BLOCO = re.compile(r"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))? (?:\[[^\]]*\] )?(?:\(v [^)]*\) )?(=.*)$")
RE_MATRIZ = re.compile(r"^([A-Z]+\d+) array (\S+) \{(=.*)\}$")
RE_EXCECAO = re.compile(r"^  except ([A-Z]+\d+) (.*)$")
RE_SUMIU = re.compile(r"^([A-Z]+\d+) no formula$")
RE_REF_QUEBRADA = re.compile(r"(?:(?:%s)!)?#REF!" % NOME_ABA)


def tokens(texto):
    # ponytail: ~4 caracteres por token (estimativa). Use tiktoken se precisar de precisao.
    return len(texto) // 4


def ref(r, c):
    return "%s%d" % (get_column_letter(c), r)


def limpa(formula):
    """Tira prefixos internos do Excel (_xlfn.XLOOKUP vira XLOOKUP) e quebras de linha (Alt+Enter)."""
    return re.sub(r"[\r\n]+", " ", re.sub(r"_xlfn\.|_xlws\.|_xlpm\.", "", formula))


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


def rotulo(formulas, valores, r, c, dr, dc):
    """Texto de cabecalho mais perto na direcao (dr, dc). Pula ate 2 vazias.
    None = nada achado; "" = achou algo que nao e rotulo (numero, formula, '%')."""
    for _ in range(3):
        r, c = r + dr, c + dc
        if r < 1 or c < 1:
            return None
        if (r, c) in formulas:
            return ""
        v = valores.get((r, c))
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        if not isinstance(v, str) or v.startswith("#") or len(re.findall(r"[^\W\d_]", v)) < 2:
            return ""
        return " ".join(re.sub(r"[\[\]]", "", v).split())[:30].strip()
    return None


def valor_curto(v):
    """Valor calculado em poucos caracteres, sem ')' (a dica '(v X)' nao entra na formula)."""
    if v is None:
        return ""
    return " ".join(fmt(v).replace(")", "").split())[:20]


def blocos(formulas, valores, vertical, buraco="blank", contexto=None, cache=None):
    """Junta celulas vizinhas com o mesmo padrao. Uma celula diferente no meio vira 'except'.
    contexto = (todas as formulas, todos os valores) da aba, so para achar rotulos."""
    todas, todos = contexto or (formulas, valores)
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
            if p0 == p1:  # celula sozinha: esquerda; acima so se a esquerda estiver vazia
                nome = rotulo(todas, todos, ini[0], ini[1], 0, -1)
                if nome is None:
                    nome = rotulo(todas, todos, ini[0], ini[1], -1, 0)
            else:
                dr, dc = (-1, 0) if vertical else (0, -1)
                nome = rotulo(todas, todos, ini[0], ini[1], dr, dc)
            if nome:
                faixa += " [%s]" % nome
            dica = valor_curto((cache or {}).get(ini)) if p0 != p1 else ""  # celula sozinha: sem dica (custa token demais)
            if dica:
                faixa += " (v %s)" % dica
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
            out[m.group(1)] = "{%s}%s" % (m.group(2), m.group(3))
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
    """{celula: formula como deve sair}; matriz leva a faixa junto, para a conferencia pegar faixa diferente."""
    return {ref(r, c): ("{%s}%s" % (f[2], limpa(f[1])) if f[0] == "a" else limpa(f[1]))
            for (r, c), f in formulas.items() if f[0] in "fa"}


def confere(linhas, meu, base=None):
    """True se as linhas expandem exatamente para as formulas originais."""
    try:
        return expandir(linhas, base) == meu
    except Exception:  # linha que o leitor nao entende = conferencia falhou
        return False


def secao_formulas(formulas, valores, buraco="blank", contexto=None, cache=None):
    if not formulas:
        return []
    return min(blocos(formulas, valores, True, buraco, contexto, cache),
               blocos(formulas, valores, False, buraco, contexto, cache),
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


MAX_REDIGIR = 10000  # caracteres que redigir() processa
MAX_MEMBRO = 5 * 2 ** 20  # bytes lidos de um membro do zip (entrada nao confiavel)
RE_URL = re.compile(r"(?!file:)[a-z][a-z0-9+.-]{1,10}://[^\s\"'<>]+", re.I)
RE_URL_ARQUIVO = re.compile(r"file:/{2,3}(?:[^\s\"'<>|;]*/)*([^\s\"'<>|;/]*)", re.I)
RE_CAMINHO_WIN = re.compile(r"(?:(?<!\w)[A-Za-z]:[\\/]|\\\\[^\\\s\"']+\\)(?:[^\\/\r\n\"'<>|;]*[\\/])*")
RE_CAMINHO_NIX = re.compile(r"(?<![\w:/.])/(?:home|Users|root|mnt|media|var|tmp|opt)/(?:[^/\s\"'<>|;]*/)*")
RE_BEARER = re.compile(r"\b(bearer|basic)\s+[^\s\"',;)\]]+", re.I)
RE_SEGREDO = re.compile(r"(\b\w{0,40}(?:password|pwd|key|token|secret|authorization)\b)\s*[=:]\s*"
                        r"(?:\"[^\"]*\"|'[^']*'|[^;,&\"')\]\r\n]+)", re.I)
RE_USERINFO = re.compile(r"(?<![\w.])[^\s/:@;\"'=]+:[^\s/@;\"']+@")  # user:senha@ sem esquema
RE_EMAIL = re.compile(r"[\w.+-]{1,64}@[\w-]+(?:\.[\w-]+)+")


def _url_limpa(m):
    try:
        u = urlsplit(m.group(0))
        return "%s://%s%s" % (u.scheme, u.netloc.rpartition("@")[2], u.path)
    except ValueError:  # URL malformada: corta no primeiro ? ou #
        return re.split(r"[?#]", m.group(0))[0].rpartition("@")[2]


def redigir(texto):
    """Tira de um texto o que identifica quem usa a planilha: so o nome do arquivo de cada caminho,
    URL sem query, fragmento e usuario, e valores depois de chaves como password/token. Redigir antes de truncar."""
    if len(texto) > MAX_REDIGIR:  # regex com retrocesso: tempo cresce com o tamanho; quem chama trunca bem antes disso
        texto = texto[:MAX_REDIGIR] + "...(cut)"
    t = RE_USERINFO.sub("", RE_URL.sub(_url_limpa, texto))
    t = RE_URL_ARQUIVO.sub(lambda m: unquote(m.group(1)), t)
    t = RE_CAMINHO_NIX.sub("", RE_CAMINHO_WIN.sub("", t))
    t = RE_SEGREDO.sub(lambda m: m.group(1) + "=***", RE_BEARER.sub(lambda m: m.group(1) + " ***", t))
    return RE_EMAIL.sub("***@***", t)


def ler_membro(zf, nome, maximo=MAX_MEMBRO):
    """Bytes de um membro do zip; ValueError se o tamanho declarado passa do teto, KeyError se nao existe."""
    if zf.getinfo(nome).file_size > maximo:
        raise ValueError("%s is larger than %d bytes" % (nome, maximo))
    return zf.read(nome)


def xml_seguro(dados):
    """Parse de XML que vem do arquivo; recusa DOCTYPE/ENTITY (entidades que explodem)."""
    if re.search(rb"<!(?:DOCTYPE|ENTITY)", dados.replace(b"\x00", b""), re.I):
        raise ValueError("XML with DOCTYPE/ENTITY refused")
    try:
        return ElementTree.fromstring(dados)
    except ElementTree.ParseError as e:
        raise ValueError("XML not readable: %s" % e)


def _injetar(buf, partes):
    """Copia o zip buf (BytesIO) para um novo, acrescentando/trocando {caminho: bytes}. So para testes."""
    buf.seek(0)
    saida = io.BytesIO()
    with zipfile.ZipFile(buf) as z, zipfile.ZipFile(saida, "w", zipfile.ZIP_DEFLATED) as o:
        for i in z.infolist():
            if i.filename not in partes:
                o.writestr(i, z.read(i.filename))
        for nome, dados in partes.items():
            o.writestr(nome, dados)
    saida.seek(0)
    return saida


ORDEM_LIVRO = ("Connections", "Queries", "External links", "Macros")


def bloco_livro(secoes):
    """Bloco '# Workbook' com as secoes do livro todo ({nome: linhas}, na ordem ORDEM_LIVRO); vazio se nenhuma tem conteudo."""
    linhas = [l for k in ORDEM_LIVRO for l in secoes.get(k, [])]
    return ["", "# Workbook"] + linhas if linhas else []


TIPO_CONEXAO = {1: "ODBC", 2: "DAO", 3: "web", 4: "text", 5: "OLE DB"}
CHAVES_CONEXAO = ("data source", "initial catalog", "provider", "driver", "database", "server")


def conexoes(fonte):
    """(linhas, avisos) das conexoes de dados (xl/connections.xml). Nunca abre a conexao.
    Da string de conexao so ficam as chaves de CHAVES_CONEXAO (lista permitida, nao lista de segredos)."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    try:
        with zipfile.ZipFile(fonte) as z:
            if "xl/connections.xml" not in z.namelist():
                return [], []
            raiz = xml_seguro(ler_membro(z, "xl/connections.xml"))
    except zipfile.BadZipFile:
        return [], []
    except (ValueError, OSError) as e:
        return [], ["- connections.xml skipped: %s" % redigir(" ".join(str(e).split()))[:100]]
    una = lambda x: " ".join(redigir(str(x)).split())
    todas = [c for c in raiz.iter() if c.tag.rpartition("}")[2] == "connection"]
    linhas, avisos, n_pq, extra = [], [], 0, 0
    for c in todas:
        filho = lambda tag: next((f for f in c if f.tag.rpartition("}")[2] == tag), None)
        db, web, txt = filho("dbPr"), filho("webPr"), filho("textPr")
        nome = una(c.get("name") or "id %s" % c.get("id", "?"))
        cadeia = (db.get("connection") if db is not None else "") or ""
        if "microsoft.mashup" in cadeia.lower():
            n_pq += 1  # uma conexao por consulta (centenas num livro grande); detalhes em ## Queries
            continue
        if len(linhas) >= LIMITE_CONEXOES:
            extra += 1
            continue
        tipo = c.get("type", "")
        partes = ["- '%s' (%s)" % (nome, TIPO_CONEXAO.get(int(tipo), "type %s" % tipo) if re.fullmatch(r"[0-9]+", tipo) else "type unknown")]
        fonte_txt = ";".join(una(kv) for kv in cadeia.split(";") if kv.split("=")[0].strip().lower() in CHAVES_CONEXAO)
        for dado in (fonte_txt, web is not None and web.get("url"), txt is not None and txt.get("sourceFile"), c.get("odcFile")):
            if dado:
                partes.append("source %s" % una(dado))
        cmd = una(db.get("command") or "") if db is not None else ""
        if cmd:
            partes.append("command %s" % (cmd[:LIMITE_COMANDO] + "...(truncated)" if len(cmd) > LIMITE_COMANDO else cmd))
        if c.get("refreshOnLoad") in ("1", "true"):
            partes.append("refreshes on open")
            avisos.append("- connection '%s' refreshes when the file opens: loaded data may differ from the saved values" % nome)
        if c.get("background") in ("1", "true"):
            partes.append("background refresh")
        linhas.append(", ".join(partes))
    if extra:
        linhas.append("(+%d more connections)" % extra)
    if n_pq:
        linhas.append("- Power Query: %d connection%s (the queries are in ## Queries)" % (n_pq, "s" if n_pq > 1 else ""))
    return linhas, avisos


RE_ITEM_XML = re.compile(r"^customXml/item\d+\.xml$")
RE_SHARED = re.compile(r'^shared\s+(#"(?:[^"]|"")*"|[\w.]+)\s*=', re.M)


def _blocos(blob):
    """(pacote zip, metadados xml) do DataMashup; ValueError se o layout nao bate (MS-QDEFF)."""
    def parte(pos):
        (n,) = struct.unpack("<I", blob[pos:pos + 4])
        if pos + 4 + n > len(blob):
            raise ValueError("part larger than blob")
        return blob[pos + 4:pos + 4 + n], pos + 4 + n
    if struct.unpack("<I", blob[:4])[0] != 0:
        raise ValueError("unknown DataMashup version")
    pacote, pos = parte(4)
    _, pos = parte(pos)  # permissoes
    meta, _ = parte(pos)
    return pacote, meta


def _destinos(meta):
    """{consulta: texto} de onde cada consulta carrega, lido dos metadados. Sem palpite: falha vira {}."""
    try:
        raiz = xml_seguro(meta[8:8 + struct.unpack("<I", meta[4:8])[0]])
        destinos = {}
        for item in raiz.iter():
            if item.tag.rpartition("}")[2] != "Item":
                continue
            caminho = next((e.text or "" for e in item.iter() if e.tag.rpartition("}")[2] == "ItemPath"), "")
            if not caminho.startswith("Section1/"):
                continue
            ent = {e.get("Type"): e.get("Value", "") for e in item.iter() if e.tag.rpartition("}")[2] == "Entry"}
            if "FillEnabled" not in ent:
                continue
            partes = ["sheet table " + ent.get("FillTarget", "s?")[1:]] if ent["FillEnabled"] == "l1" else ["connection only"]
            if "l1" in (ent.get("LoadedToAnalysisServices"), ent.get("FillToDataModelEnabled")):
                partes.append("data model")
            destinos[unquote(caminho[len("Section1/"):])] = ", ".join(partes)
        return destinos
    except Exception:
        return {}


def consultas(fonte):
    """([(nome, codigo M, destino)], legivel). ([], True) se nao ha Power Query; legivel False se ha mas nao deu para ler.
    Nunca executa nada: so desempacota base64, o zip interno e Section1.m."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    try:
        with zipfile.ZipFile(fonte) as z:
            for nome in sorted(n for n in z.namelist() if RE_ITEM_XML.match(n)):
                try:
                    bruto = ler_membro(z, nome)
                except Exception:  # grande demais, cifrado, compressao estranha: nao da para saber se e DataMashup
                    continue
                if b"DataMashup" not in bruto.replace(b"\x00", b""):
                    continue
                try:
                    raiz = xml_seguro(bruto)
                    if raiz.tag.rpartition("}")[2] != "DataMashup":
                        continue
                    pacote, meta = _blocos(base64.b64decode("".join((raiz.text or "").split())))
                    with zipfile.ZipFile(io.BytesIO(pacote)) as zi:
                        m = ler_membro(zi, "Formulas/Section1.m")
                    m = m.decode("utf-16" if m[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig", "replace")
                except Exception:
                    return [], False
                destinos = _destinos(meta)
                achadas = list(RE_SHARED.finditer(m))
                lista = []
                for i, a in enumerate(achadas):
                    fim = achadas[i + 1].start() if i + 1 < len(achadas) else len(m)
                    nome_q = a.group(1)
                    nome_q = nome_q[2:-1].replace('""', '"') if nome_q.startswith('#"') else nome_q
                    lista.append((nome_q, m[a.end():fim].strip().rstrip(";").strip(), destinos.get(nome_q, "")))
                return lista, True
    except zipfile.BadZipFile:
        pass
    except (ValueError, OSError):
        return [], False
    return [], True


def secao_consultas(fonte):
    """Linhas do ## Queries. O codigo M vai como escrito (redigido, depois truncado), recuado para nao virar cabecalho."""
    lista, legivel = consultas(fonte)
    if not legivel:
        return ["Queries present but unreadable"]
    linhas = []
    for nome, codigo, destino in lista[:LIMITE_CONSULTAS]:
        linhas.append("- %s%s" % (" ".join(redigir(nome).split()), ", loads to: " + destino if destino else ""))
        codigo = redigir(codigo)
        if len(codigo) > LIMITE_M:
            codigo = codigo[:LIMITE_M] + "...(truncated)"
        linhas += ["    " + l for l in codigo.splitlines()]
    if len(lista) > LIMITE_CONSULTAS:
        linhas.append("(+%d more queries)" % (len(lista) - LIMITE_CONSULTAS))
    return linhas


CFB_ASSINATURA = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
CFB_FIM = 0xFFFFFFFE
MAX_FLUXO = 5 * 2 ** 20   # bytes de um stream do container
MAX_FLUXOS = 20 * 2 ** 20  # bytes de todos os streams juntos
MAX_VBA_MODULO = 2 ** 20   # bytes descompactados de um modulo VBA
MAX_DESCOMPACTADO = 8 * 2 ** 20  # idem, todos os modulos


def cfb_streams(d):
    """{caminho: bytes} dos streams de um container OLE/CFB (formato do vbaProject.bin). ValueError se invalido.
    Entrada nao confiavel: toda cadeia tem teto e deteccao de ciclo, streams grandes demais sao pulados.
    # ponytail: sem cadeia DIFAT (FAT acima de 109 setores, ~7 MB com setor de 512); reconstruir se um arquivo real pedir."""
    if d[:8] != CFB_ASSINATURA or len(d) < 512:
        raise ValueError("not a CFB container")
    sh, msh = struct.unpack_from("<HH", d, 30)
    if sh not in (9, 12) or msh != 6:
        raise ValueError("unsupported sector size")
    ss, ms = 1 << sh, 64
    nfat, dir0, _, corte, mini0, nmini, _, ndifat = struct.unpack_from("<IIIIIIII", d, 44)
    if ndifat or nfat > 109:
        raise ValueError("DIFAT chain not supported")
    limite = len(d) // ss + 1

    def setor(i):
        if (i + 2) * ss > len(d):
            raise ValueError("sector beyond end of file")
        return d[(i + 1) * ss:(i + 2) * ss]

    def cadeia(ini, tabela, leitor, n=None):  # n: so os primeiros n setores (o stream declara quanto precisa)
        partes, visto, i = [], set(), ini
        while i != CFB_FIM and (n is None or len(partes) < n):
            if i in visto or i >= len(tabela) or len(partes) > limite:
                raise ValueError("bad sector chain")
            visto.add(i)
            partes.append(leitor(i))
            i = tabela[i]
        return b"".join(partes)

    fat = []
    for i in struct.unpack_from("<109I", d, 76)[:nfat]:
        fat += struct.unpack("<%dI" % (ss // 4), setor(i))
    diretorio = cadeia(dir0, fat, setor)
    ent = []
    for o in range(0, min(len(diretorio), 4096 * 128), 128):
        nome, nlen, tipo = struct.unpack_from("<64sHB", diretorio, o)
        esq, dir_, filho = struct.unpack_from("<III", diretorio, o + 68)
        ini, tam = struct.unpack_from("<IQ", diretorio, o + 116)
        ent.append((nome[:max(nlen - 2, 0)].decode("utf-16-le", "replace"), tipo, esq, dir_, filho, ini, tam & 0xFFFFFFFF))
    if not ent:
        raise ValueError("empty directory")
    bruto = cadeia(mini0, fat, setor) if nmini else b""
    minifat = struct.unpack("<%dI" % (len(bruto) // 4), bruto)
    ministream = cadeia(ent[0][5], fat, setor)[:ent[0][6]] if ent[0][5] != CFB_FIM else b""

    saida, visto, total, pilha = {}, set(), 0, [(ent[0][4], "")]
    while pilha:
        j, pref = pilha.pop()
        if j == 0xFFFFFFFF:
            continue
        if j in visto or j >= len(ent):
            raise ValueError("bad directory tree")
        visto.add(j)
        nome, tipo, esq, dir_, filho, ini, tam = ent[j]
        pilha += [(esq, pref), (dir_, pref)]
        if tipo == 1:
            pilha.append((filho, pref + nome + "/"))
        elif tipo == 2 and tam <= MAX_FLUXO:
            total += tam
            if total > MAX_FLUXOS:
                raise ValueError("streams too large")
            if not tam:
                dados = b""
            elif tam < corte:
                dados = cadeia(ini, minifat, lambda i: ministream[i * ms:(i + 1) * ms], -(-tam // ms))
            else:
                dados = cadeia(ini, fat, setor, -(-tam // ss))
            if len(dados) < tam:
                raise ValueError("stream shorter than declared")
            saida[pref + nome] = dados[:tam]
    return saida


def ovba_descomprimir(d, maximo=MAX_FLUXO):
    """Descompacta um stream VBA (MS-OVBA 2.4.1). ValueError se corrompido ou grande demais."""
    if not d or d[0] != 1:
        raise ValueError("bad compression signature")
    saida, pos = bytearray(), 1
    while pos + 2 <= len(d):
        cab = d[pos] | d[pos + 1] << 8
        fim, comprimido, ini = min(pos + (cab & 0xFFF) + 3, len(d)), cab >> 15, len(saida)
        if cab >> 12 & 7 != 3:
            raise ValueError("bad chunk signature")
        pos += 2
        if not comprimido:
            saida += d[pos:pos + 4096]
            pos = fim
            continue
        while pos < fim:
            marcas = d[pos]
            pos += 1
            for bit in range(8):
                if pos >= fim:
                    break
                if not marcas >> bit & 1:
                    saida.append(d[pos])
                    pos += 1
                    continue
                if pos + 2 > fim or len(saida) == ini:
                    raise ValueError("bad copy token")
                tok = d[pos] | d[pos + 1] << 8
                pos += 2
                bits = 4
                while 1 << bits < len(saida) - ini:
                    bits += 1
                desloc, tam = (tok >> (16 - bits)) + 1, (tok & (0xFFFF >> bits)) + 3
                if desloc > len(saida) - ini:
                    raise ValueError("bad copy offset")
                for _ in range(tam):
                    saida.append(saida[-desloc])
        if len(saida) > maximo:
            raise ValueError("decompressed stream too large")
    return bytes(saida)


def _texto_vba(b, cp):
    try:
        return b.decode(cp)
    except (LookupError, UnicodeDecodeError):
        return b.decode("latin-1")


def _decifra_projeto(hexa):
    """Dado de CMG=/DPB= do stream PROJECT (MS-OVBA 2.4.3.2, cifra com semente). ValueError se nao der."""
    enc = bytes(bytearray.fromhex(hexa))
    if len(enc) < 8:
        raise ValueError("too short")
    u1, e1, e2, saida = enc[0] ^ enc[2], enc[2], enc[1], bytearray()
    for b in enc[3:]:
        d = b ^ ((e2 + u1) & 0xFF)
        e2, e1, u1 = e1, b, d
        saida.append(d)
    saida = saida[(enc[0] & 6) >> 1:]
    return bytes(saida[4:4 + struct.unpack("<I", saida[:4])[0]])


def projeto_protegido(projeto):
    """True se o projeto VBA tem senha ou protecao de visualizacao. Excel escreve CMG/DPB em todo projeto
    (sem senha eles decifram para 0), entao a mera presenca nao basta. Se existem mas nao decifram, conta como protegido."""
    try:
        cmg, dpb = (re.search(r'^%s="([0-9A-Fa-f]+)"' % k, projeto, re.M) for k in ("CMG", "DPB"))
        return bool(cmg and any(_decifra_projeto(cmg.group(1)))) or bool(dpb and _decifra_projeto(dpb.group(1)) != b"\x00")
    except (ValueError, IndexError, struct.error):
        return True


def vba_modulos(dir_bytes):
    """[{nome, fluxo, offset, tipo}] do stream 'dir' ja descompactado (MS-OVBA 2.3.4.2): registros id(2) tamanho(4) dado."""
    pos, cp, mods, atual = 0, "cp1252", [], None
    while pos + 6 <= len(dir_bytes):
        rid, n = struct.unpack_from("<HI", dir_bytes, pos)
        dado = dir_bytes[pos + 6:pos + 6 + n]
        pos += 6 + n + (2 if rid == 9 else 0)  # versao do projeto: 2 bytes alem do tamanho declarado
        if rid == 3 and len(dado) == 2:
            cp = "cp%d" % struct.unpack("<H", dado)[0]
        elif rid == 0x19:
            atual = {"nome": _texto_vba(dado, cp), "fluxo": _texto_vba(dado, cp), "offset": 0, "tipo": ""}
            mods.append(atual)
        elif atual is not None and rid == 0x1A:
            atual["fluxo"] = _texto_vba(dado, cp)
        elif atual is not None and rid == 0x31 and len(dado) == 4:
            atual["offset"] = struct.unpack("<I", dado)[0]
        elif atual is not None and rid in (0x21, 0x22):
            atual["tipo"] = "module" if rid == 0x21 else "class or document"
        elif rid == 0x2B:
            atual = None
        elif rid == 0x10:
            break
    return mods


RE_PROCEDIMENTO = re.compile(r"^[ \t]*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.I | re.M)
RE_ESCRITA = re.compile(r"^[ \t]*((?:\w+(?:\([^)\r\n]*\))?\.)*(?:(?:Range|Cells)\s*\(.*\)(?:\.\w+)*|\.(?:Value2?|Formula\w*)))\s*=(?!=)", re.I | re.M)


def secao_macros(fonte, com_fonte=False):
    """Linhas do ## Macros ([] se nao ha VBA ou nada deu para ler). Nunca executa nada.
    Sem com_fonte so os nomes dos modulos; com ela tipo, procedimentos, escritas em celulas e o codigo (redigido e truncado)."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    try:
        with zipfile.ZipFile(fonte) as z:
            if "xl/vbaProject.bin" not in z.namelist():
                return []
            fluxos = cfb_streams(ler_membro(z, "xl/vbaProject.bin"))
    except Exception:  # container ruim ou grande demais: fica so o aviso antigo
        return []
    nomes = sorted(n[4:] for n in fluxos if n.startswith("VBA/") and n[4:] not in ("dir", "_VBA_PROJECT") and not n[4:].startswith("__SRP_"))
    projeto = fluxos.get("PROJECT", b"").decode("latin-1")
    protegido = projeto_protegido(projeto)
    tipos = {}
    for chave, tipo in (("Module", "module"), ("Class", "class"), ("Document", "document"), ("BaseClass", "form")):
        for nome in re.findall(r"^%s=([^/\r\n]+)" % chave, projeto, re.M):
            tipos[nome] = tipo
    limpa1 = lambda x: " ".join(redigir(x).split())
    linhas = ["- VBA project is password protected: source never shown"] if protegido else []
    if not com_fonte or protegido:
        if nomes:
            linhas.append("modules: " + ", ".join(limpa1(n) for n in nomes[:LIMITE_PIVOTS]) + (" (+%d more)" % (len(nomes) - LIMITE_PIVOTS) if len(nomes) > LIMITE_PIVOTS else ""))
        if nomes and not protegido:
            linhas.append("(source not shown: run with --macros-fonte)")
        return linhas
    try:
        mods = vba_modulos(ovba_descomprimir(fluxos["VBA/dir"]))
    except Exception:
        return linhas + ["modules: " + ", ".join(limpa1(n) for n in nomes[:LIMITE_PIVOTS]), "(project index unreadable: no source)"]
    usado = gasto = 0  # gasto: bytes descompactados (modulos podem apontar para o mesmo fluxo bomba)
    for m in mods[:LIMITE_PIVOTS]:
        nome = limpa1(m["nome"])
        tipo = tipos.get(m["nome"]) or m["tipo"]
        try:
            if gasto >= MAX_DESCOMPACTADO:
                raise ValueError("decompression budget used up")
            gasto += MAX_VBA_MODULO  # pior caso; o real e devolvido abaixo
            bruto = ovba_descomprimir(fluxos["VBA/" + m["fluxo"]][m["offset"]:], MAX_VBA_MODULO)
            gasto += len(bruto) - MAX_VBA_MODULO
            codigo = _texto_vba(bruto, "cp1252")[:MAX_VBA_MODULO]
        except Exception:
            linhas.append("- %s%s: source unreadable" % (nome, " (%s)" % tipo if tipo else ""))
            continue
        codigo = "\n".join(l for l in codigo.splitlines() if not l.startswith("Attribute VB_"))
        procs = ["%s %s" % (a.split()[0], b) for a, b in RE_PROCEDIMENTO.findall(codigo)]
        linhas.append("- %s%s%s" % (nome, " (%s)" % tipo if tipo else "", ": " + ", ".join(procs) if procs else ""))
        escritas = list(dict.fromkeys(limpa1(x.group(1))[:120] for x in RE_ESCRITA.finditer(codigo)))  # so o destino, sem repetir
        if escritas:
            linhas.append("  writes to cells:")
            linhas += ["    " + e for e in escritas[:LIMITE_ESCRITAS]]
            if len(escritas) > LIMITE_ESCRITAS:
                linhas.append("    (+%d more)" % (len(escritas) - LIMITE_ESCRITAS))
        codigo = redigir(codigo.strip())
        if not codigo:
            continue
        if usado >= LIMITE_VBA_TOTAL:
            linhas.append("  (source omitted: total cap of %d characters reached)" % LIMITE_VBA_TOTAL)
            continue
        codigo = codigo[:min(LIMITE_VBA, LIMITE_VBA_TOTAL - usado)] + ("...(truncated)" if len(codigo) > min(LIMITE_VBA, LIMITE_VBA_TOTAL - usado) else "")
        usado += len(codigo)
        linhas += ["  source:"] + ["    " + l for l in codigo.splitlines()]
    if len(mods) > LIMITE_PIVOTS:
        linhas.append("(+%d more modules)" % (len(mods) - LIMITE_PIVOTS))
    return linhas


def links_externos(wb, textos):
    """(linhas, avisos) dos links para outras planilhas, na ordem de [N] nas formulas. Nunca abre o arquivo ligado.
    textos = todas as formulas do livro, para contar o uso de cada [N]. wb._external_links nao e API publica."""
    # ponytail: "[1]x" dentro de texto entre aspas tambem conta; raro, so afeta o numero
    uso = Counter(int(i) for t in textos for i in set(re.findall(r"\[(\d+)\](?=[\w'])", t)))
    linhas, avisos = [], []
    for n, link in enumerate(getattr(wb, "_external_links", None) or [], 1):
        try:
            alvo = redigir(str(link.file_link.Target))
            if "://" not in alvo:
                alvo = re.split(r"[\\/]", re.split(r"[?#]", alvo)[0])[-1]  # so o nome do arquivo
            alvo = " ".join(alvo.split()) or "unknown target"
            if "://" in alvo:
                avisos.append("- external link [%d] points to a URL: %s" % (n, alvo))
            livro = link.externalBook
            abas = [" ".join(str(a).split()) for a in livro.sheetNames.sheetName] if livro.sheetNames else []
            if livro.sheetDataSet is None:
                avisos.append("- external link [%d] has no saved values: formula results that use it cannot be trusted" % n)
        except Exception:
            alvo, abas = "unknown target", []
        linhas.append("[%d] %s%s, %d formulas" % (n, alvo, ", sheets: " + ", ".join(abas) if abas else "", uso[n]))
    return linhas, avisos


def nomes_definidos(wb):
    return ["%s = %s" % (d.name, d.attr_text) for d in _nomes(wb)]


def _nomes(wb, globais=False):
    dn = wb.defined_names
    lista = list(dn.values()) if hasattr(dn, "values") else list(getattr(dn, "definedName", []))
    for ws in ([] if globais else wb.worksheets):  # nomes com escopo de aba (openpyxl 3.1+)
        lista.extend(getattr(ws, "defined_names", {}).values())
    return lista


RE_TABELA = re.compile(r"^([^\[\]]*)\[(.*)\]$")


def parte_aba(texto):
    """"'P 1'!A1" -> ("P 1", "A1"); "A1" -> (None, "A1")."""
    aba, sep, resto = texto.rpartition("!")
    if not sep:
        return None, texto
    if len(aba) > 1 and aba[0] == aba[-1] == "'":
        aba = aba[1:-1].replace("''", "'")
    return aba, resto


def acha_alvo(wb, texto):
    """'Aba'!B5 / Aba!B5 / B5 (aba ativa) -> (titulo, linha, coluna). ValueError se nao der."""
    aba, cel = parte_aba(texto.strip())
    m = PECA_CELULA.match(cel.upper())
    if not m or column_index_from_string(m.group(2)) > 16384:
        raise ValueError("celula invalida: %s (use Aba!B5, uma celula so)" % texto)
    nomes = {ws.title.lower(): ws.title for ws in wb.worksheets}
    titulo = wb.active.title if aba is None else nomes.get(aba.lower())
    if titulo is None:
        raise ValueError("aba nao encontrada: %s" % aba)
    return titulo, int(m.group(4)), column_index_from_string(m.group(2))


def faixa(texto):
    """'A1:B9', 'A:A', '2:2', 'B5' -> (r1, c1, r2, c2), None = sem limite. None se nao for faixa."""
    partes = texto.upper().split(":")
    if len(partes) == 1:  # sozinho so vale celula: "TIR" e nome, nao a coluna TIR
        if not PECA_CELULA.match(partes[0]):
            return None
        partes = partes * 2
    if len(partes) != 2:
        return None
    (a, b) = partes
    ma, mb = PECA_CELULA.match(a), PECA_CELULA.match(b)
    if ma and mb:
        r1, r2 = sorted((int(ma.group(4)), int(mb.group(4))))
        c1, c2 = sorted((column_index_from_string(ma.group(2)), column_index_from_string(mb.group(2))))
        return r1, c1, r2, c2
    ma, mb = PECA_COLUNA.match(a), PECA_COLUNA.match(b)
    if ma and mb:
        c1, c2 = sorted((column_index_from_string(ma.group(2)), column_index_from_string(mb.group(2))))
        return None, c1, None, c2
    ma, mb = PECA_LINHA.match(a), PECA_LINHA.match(b)
    if ma and mb:
        r1, r2 = sorted((int(ma.group(2)), int(mb.group(2))))
        return r1, None, r2, None
    return None


def precedentes(wb, dados, alvo):
    """Celulas de que alvo depende, recursivo entre abas.
    dados = {aba: (formulas, valores)}. Retorna ({aba: set(pos)}, avisos, nomes usados)."""
    abas = {t.lower(): t for t in dados}
    indice, matrizes, tabelas = {}, defaultdict(list), {}
    for aba, (formulas, valores) in dados.items():
        cols = defaultdict(list)
        for (r, c) in set(formulas) | set(valores):
            cols[c].append(r)
        indice[aba] = {c: sorted(rs) for c, rs in cols.items()}
        for (r, c), f in formulas.items():
            if f[0] == "a":
                r1, c1, r2, c2 = faixa(f[2])
                matrizes[aba].append((r1, c1, r2, c2, (r, c)))
    for ws in wb.worksheets:
        for t in ws.tables.values():
            r1, c1, r2, c2 = faixa(t.ref)
            cab = 1 if t.headerRowCount is None else t.headerRowCount
            tabelas[t.displayName.lower()] = (ws.title, r1 + cab, c1, r2, c2,
                                              [col.name.lower() for col in t.tableColumns])
    nomes = {(None, d.name.lower()): (d.name, d.attr_text) for d in _nomes(wb, globais=True)}
    for ws in wb.worksheets:  # nome com escopo de aba ganha do global nessa aba
        for d in getattr(ws, "defined_names", {}).values():
            nomes[(ws.title, d.name.lower())] = (d.name, d.attr_text)
    avisos, exemplo, usados = Counter(), {}, set()

    def celulas(aba, r1, c1, r2, c2):
        for c, linhas in indice[aba].items():
            if c1 is not None and not c1 <= c <= c2:
                continue
            ini = 0 if r1 is None else bisect.bisect_left(linhas, r1)
            fim = len(linhas) if r2 is None else bisect.bisect_right(linhas, r2)
            for r in linhas[ini:fim]:
                yield (aba, r, c)
        for a1, b1, a2, b2, ancora in matrizes[aba]:
            if (r1 is None or (a1 <= r2 and r1 <= a2)) and (c1 is None or (b1 <= c2 and c1 <= b2)):
                yield (aba, ancora[0], ancora[1])

    def tabela(nome, dentro, aqui, pos):
        if nome:
            t = tabelas.get(nome.lower())
        else:  # [@Col]: tabela que contem a propria celula
            t = next((t for t in tabelas.values() if t[0] == aqui
                      and t[1] - 1 <= pos[0] <= t[3] and t[2] <= pos[1] <= t[4]), None)
        if t is None:
            return None
        aba, r1, c1, r2, c2, cols = t
        pedidas = [x.strip().lstrip("@").lower() for x in re.split(r"[\[\],:]", dentro)]
        achadas = [c1 + cols.index(x) for x in pedidas if x in cols]
        if achadas:
            c1, c2 = min(achadas), max(achadas)
        return list(celulas(aba, r1, c1, r2, c2))

    def resolve(v, aqui, pos, abertos=()):
        """Operando de formula -> lista de celulas, ou None se nao der para seguir."""
        if re.match(r"^'?\[\d+\]", v):
            avisos[(aqui, "external file reference (not followed)")] += 1
            return []
        m = RE_TABELA.match(v)
        if m and "!" not in v:
            return tabela(m.group(1), m.group(2), aqui, pos)
        aba, alvo_txt = parte_aba(v)
        f = faixa(alvo_txt)
        if f is not None:
            aba = aqui if aba is None else abas.get(aba.lower())
            return None if aba is None else list(celulas(aba, *f))
        achado = nomes.get((aqui, v.lower())) or nomes.get((None, v.lower())) if aba is None else None
        if achado:
            nome, destino = achado
            usados.add(nome.lower())
            if "(" in destino or v.lower() in abertos:  # formula ou nome circular
                return None
            if re.match(r'^=?\s*(-?[\d.]+(E[+-]?\d+)?%?|".*"|TRUE|FALSE)\s*$', destino, re.I):
                return []  # constante (Taxa = 0.05): nada para seguir
            saida = []
            for parte in destino.split(","):
                achou = resolve(parte.strip(), aqui, pos, abertos + (v.lower(),))
                if achou is None:
                    return None
                saida += achou
            return saida
        if aba is None and v.lower() in tabelas:
            return tabela(v, "", aqui, pos)
        return None

    memo = {}
    fila, vistos = list(celulas(*alvo[:1] + (alvo[1], alvo[2], alvo[1], alvo[2]))) or [alvo], set()
    while fila:
        k = fila.pop()
        if k in vistos:
            continue
        vistos.add(k)
        f = dados[k[0]][0].get(k[1:])
        if not f or f[0] == "d":
            continue
        try:
            itens = Tokenizer(f[1]).items
        except Exception:  # ex.: B1# (spill) o Tokenizer nao entende
            chave = (k[0], "formula not parsed, refs not followed")
            avisos[chave] += 1
            exemplo.setdefault(chave, ref(*k[1:]))
            continue
        for t in itens:
            if t.type == Token.OPERAND and t.subtype == Token.RANGE:
                chave_memo = (t.value, k[0])
                if "[" in t.value:  # [@Col] depende da linha: sem cache
                    achou = resolve(t.value, k[0], k[1:])
                elif chave_memo in memo:
                    achou = memo[chave_memo]
                else:
                    achou = memo[chave_memo] = resolve(t.value, k[0], k[1:])
                if achou is None:
                    chave = (k[0], "unresolved reference %s" % t.value)
                    avisos[chave] += 1
                    exemplo.setdefault(chave, ref(*k[1:]))
                else:
                    fila.extend(x for x in achou if x not in vistos)
    cadeia = defaultdict(set)
    for aba, r, c in vistos:
        cadeia[aba].add((r, c))
    linhas = ["- '%s': %s%s%s" % (aba, txt, " in %s" % exemplo[(aba, txt)] if (aba, txt) in exemplo else "",
                                   " (x%d)" % n if n > 1 else "") for (aba, txt), n in sorted(avisos.items())]
    return cadeia, linhas, usados


def ler_abas(wb, wb_v):
    """[(aba, formulas, valores, cache)], tamanho do dump celula a celula, True se nenhuma formula tem valor salvo."""
    abas, bruto_chars, sem_cache = [], 0, True
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
        abas.append((ws, formulas, valores, cache))

    return abas, bruto_chars, sem_cache


def tipo_formato(nf):
    """percent / date / currency a partir do formato do numero; "" para o resto (General, 0.00, texto...)."""
    if "%" in nf:
        return "percent"
    if is_date_format(nf):
        return "date"
    return "currency" if re.search(r"[$\u20ac\u00a3]", nf) else ""


def texto_nota(comentario):
    """Texto do comentario. Comentario encadeado vem com aviso fixo do Excel antes do texto; tira o aviso."""
    t = re.sub(r"^\[(?:Coment\u00e1rio encadeado|Threaded comment)\].*?(?:Coment\u00e1rio|Comment):", "", comentario.text, flags=re.S)
    return " ".join(t.split())


def secao_extras(ws, so_estas=None):
    """(linhas de formatos, linhas de notas) de uma aba. Fatos lidos do arquivo, sem palpite."""
    por_tipo, notas = defaultdict(list), []
    for linha in ws.iter_rows():
        for c in linha:
            if c.value is None or (so_estas is not None and (c.row, c.column) not in so_estas):
                continue
            t = tipo_formato(c.number_format)
            if t:
                por_tipo[t].append((c.column, c.row))
            if c.comment and texto_nota(c.comment):
                notas.append("%s note: %s" % (ref(c.row, c.column), texto_nota(c.comment)[:LIMITE_NOTA]))
    formatos = []
    for t, cels in sorted(por_tipo.items()):
        faixas, ini = [], None
        for col, r in sorted(cels):  # colunas seguidas viram faixas verticais
            if ini and ini[0] == col and r == fim + 1:
                fim = r
                continue
            if ini:
                faixas.append((ini, fim))
            ini, fim = (col, r), r
        faixas.append((ini, fim))
        txt = [ref(i[1], i[0]) + (":" + ref(f, i[0]) if f != i[1] else "") for i, f in faixas]
        # ponytail: so as primeiras faixas; se isso nao bastar, agrupar colunas vizinhas
        formatos.append("%s: %s%s" % (t, ", ".join(txt[:LIMITE_FAIXAS]),
                                       " (+%d more)" % (len(txt) - LIMITE_FAIXAS) if len(txt) > LIMITE_FAIXAS else ""))
    if len(notas) > LIMITE_NOTAS:
        notas = notas[:LIMITE_NOTAS] + ["(+%d more notes)" % (len(notas) - LIMITE_NOTAS)]
    return formatos, notas


def _pivot(p):
    """(linhas, avisos) de uma tabela dinamica. Campos vem do cache pelo indice."""
    una = lambda x: " ".join(str(x).split())
    nome = una(p.name or "?")
    cache = getattr(p, "cache", None)
    campos = [f.name for f in cache.cacheFields] if cache is not None else []

    def campo(i):
        if i == -2:
            return "Values"
        return una(campos[i]) if 0 <= i < len(campos) and campos[i] else "field %s" % i

    fonte, avisos = "source unknown", []
    nao_le = "pivot '%s' source could not be read" % nome
    src = getattr(cache, "cacheSource", None)
    if src is None:
        avisos.append("pivot '%s' has no readable cache, source unknown" % nome)
    elif src.type == "worksheet" and src.worksheetSource is not None:
        w = src.worksheetSource
        if w.name:
            fonte = "table or name %s" % una(w.name)
        elif w.sheet and w.ref:
            fonte = "'%s'!%s" % (una(w.sheet), una(w.ref))
        else:
            avisos.append(nao_le)
    elif src.type and src.type != "worksheet":
        fonte = "%s (not a worksheet range)" % src.type
    else:
        avisos.append(nao_le)
    if cache is not None and cache.refreshOnLoad:
        avisos.append("pivot '%s' refreshes when the file opens: its numbers may differ from what the file shows" % nome)
    linhas = ["- '%s' at %s, source %s" % (nome, una(p.location.ref) if p.location is not None else "?", fonte)]
    for papel, itens in (("rows", [campo(x.x) for x in p.rowFields]),
                         ("columns", [campo(x.x) for x in p.colFields]),
                         ("filters", [campo(x.fld) for x in p.pageFields]),
                         ("values", ["%s = %s(%s)" % (una(d.name or "?"), d.subtotal, campo(d.fld)) for d in p.dataFields])):
        if itens:
            linhas.append("  %s: %s" % (papel, ", ".join(itens)))
    return linhas, avisos


def secao_pivots(ws):
    """(linhas, avisos) das tabelas dinamicas de uma aba. Fatos lidos do arquivo, sem palpite;
    ws._pivots nao e API publica, entao qualquer erro vira um aviso."""
    pivots = list(getattr(ws, "_pivots", None) or [])
    linhas, avisos = [], []
    for p in pivots[:LIMITE_PIVOTS]:
        try:
            l, a = _pivot(p)
        except Exception:
            l, a = ["- pivot unreadable"], ["a pivot could not be read"]
        linhas += l
        avisos += ["- '%s': %s" % (ws.title, x) for x in a]
    if len(pivots) > LIMITE_PIVOTS:
        linhas.append("(+%d more pivots)" % (len(pivots) - LIMITE_PIVOTS))
    return linhas, avisos


def converter(fonte, alvo=None, macros_fonte=False):
    """Retorna (texto, numeros). fonte = caminho ou arquivo aberto.
    alvo = "Aba!B5": so a cadeia de precedentes dessa celula. macros_fonte: mostra o codigo VBA (## Macros)."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    wb = load_workbook(fonte)
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    wb_v = load_workbook(fonte, read_only=True, data_only=True)

    avisos, mapa, corpo = [], [], []
    falhas, total_formulas = [], 0
    ja_vistas = []  # (nome, esperado) das abas anteriores
    s_form = s_dados = 0
    ha_pivots = False
    secoes_livro = {}  # {nome em ORDEM_LIVRO: linhas}; emitidas uma vez no fim, so fora do --celula
    todos_textos = []

    abas, bruto_chars, sem_cache = ler_abas(wb, wb_v)

    cadeia, contexto, todas = None, None, 0
    if alvo is not None:
        alvo_pos = acha_alvo(wb, alvo)
        todas = sum(1 for _, fs, _, _ in abas for f in fs.values() if f[0] in "fa")
        cadeia, avisos_cadeia, usados = precedentes(wb, {ws.title: (f, v) for ws, f, v, _ in abas}, alvo_pos)
    for ws, formulas, valores, cache in abas:
        if cadeia is not None:
            meus = cadeia.get(ws.title)
            if not meus:
                continue
            contexto = (formulas, valores)  # rotulos olham a aba inteira
            formulas = {p: f for p, f in formulas.items() if p in meus}
            valores = {p: v for p, v in valores.items() if p in meus}

        textos = [f[1] for f in formulas.values() if f[0] in "fa"]
        total_formulas += len(textos)
        if cadeia is None:
            todos_textos += textos
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
        linhas_f = secao_formulas(formulas, valores, "blank" if cadeia is None else "not in chain", contexto, cache)
        if not confere(linhas_f, meu):
            falhas.append(ws.title)
            linhas_f = por_celula(formulas)
        for base_nome, base in ja_vistas:
            iguais = sum(1 for k, v in meu.items() if base.get(k) == v)
            uniao = len(set(meu) | set(base))
            if uniao and iguais / float(uniao) >= RATIO_DUPLICADA:
                dif = diferenca(formulas, meu, base_nome, base)
                if len("\n".join(dif)) < len("\n".join(linhas_f)) and confere(dif, meu, base):
                    linhas_f = dif
                    break
        ja_vistas.append((ws.title, meu))

        linhas_d = secao_dados(valores)
        formatos, notas = secao_extras(ws, meus if cadeia is not None else None)
        pivots, avisos_pivots = secao_pivots(ws) if cadeia is None else ([], [])
        avisos += avisos_pivots
        bloco = ["", "# Sheet %s" % nome]
        if linhas_f:
            bloco += ["## Formulas"] + linhas_f
        if linhas_d:
            bloco += ["## Data"] + linhas_d
        if formatos:
            bloco += ["## Formats"] + formatos
        if notas:
            bloco += ["## Notes"] + notas
        if pivots:
            bloco += ["## Pivots"] + pivots
            ha_pivots = True
        s_form += tokens("\n".join(linhas_f))
        s_dados += tokens("\n".join(linhas_d))
        corpo += bloco

    if cadeia is None:
        macros = secao_macros(fonte, macros_fonte)
        if macros:
            secoes_livro["Macros"] = ["## Macros"] + macros
        queries = secao_consultas(fonte)
        if queries:
            secoes_livro["Queries"] = ["## Queries"] + queries
        con, avisos_con = conexoes(fonte)
        if con or avisos_con:
            secoes_livro["Connections"] = ["## Connections"] + con if con else []
            avisos += avisos_con
        links, avisos_links = links_externos(wb, todos_textos)
        if links:
            secoes_livro["External links"] = ["## External links"] + links
            avisos += avisos_links
        corpo += bloco_livro(secoes_livro)
    wb_v.close()  # read_only segura o arquivo aberto
    if tem_vba(fonte):
        avisos.insert(0, "- File has VBA macros: some values may be written by a macro, not by formulas.")
    if total_formulas and sem_cache:
        avisos.insert(0, "- No calculated values saved in the file (never opened in Excel?): error checks are blind.")
    if falhas:
        avisos.insert(0, "- WARNING: compact view failed the self-check on %d sheet(s): %s. "
                         "Those sheets list one formula per cell." % (len(falhas), ", ".join(falhas)))
    nomes = nomes_definidos(wb)
    if cadeia is not None:
        avisos += avisos_cadeia
        nomes = [n for n in nomes if n.split(" = ")[0].lower() in usados]
    if nomes:
        mapa += ["Defined names:"] + ["  " + n for n in nomes]

    cabeca = [
        "# excel_ai dump",
        "How to read formulas: '<range> <formula>' gives the formula of the first cell;",
        "the rest of the range is the same formula copied (Excel fill), refs shifting as usual.",
        "'[label]' after the range is the nearest text header (above a column, left of a row): a hint, not part of the formula.",
        "'(v X)' is the value Excel last calculated for the first cell of the range: a hint, not part of the formula.",
        "'except <cell> ...' overrides one cell inside the range above it.",
        "'same formulas as X' means the sheet is a copy of X except the listed cells.",
        "'## Formats' lists cells shown as percent/date/currency in Excel; '## Notes' has the cell comments (author text).",
    ]
    if ha_pivots:
        cabeca.append("'## Pivots' lists pivot tables: where each sits, what feeds it, and which fields are rows, columns, filters and values.")
    if "Connections" in secoes_livro:
        cabeca.append("'## Connections' (after the last sheet) lists data connections (database, web, text file): sheets fed by them hold loaded data that may be stale. Credentials are never shown.")
    if "Macros" in secoes_livro:
        cabeca.append("'## Macros' (after the last sheet) lists the VBA modules; code is shown as written, never run, and only with --macros-fonte.")
    if "Queries" in secoes_livro:
        cabeca.append("'## Queries' (after the last sheet) has the Power Query M code as written, never run; review it before sharing.")
    if "External links" in secoes_livro:
        cabeca.append("'## External links' (after the last sheet) maps each [N] in formulas to the workbook it points to; the linked file is never opened.")
    if cadeia is not None:
        aba, r, c = alvo_pos
        valor = next(cc for w, _, _, cc in abas if w.title == aba).get((r, c))
        cabeca += [
            "Precedents of '%s'!%s (cached value: %s): only the %d cells it depends on, in %d sheet(s)%s."
            % (aba, ref(r, c), "none" if valor is None else fmt(valor), sum(map(len, cadeia.values())), len(cadeia),
               "; some refs could not be followed, see Warnings" if avisos_cadeia else ""),
            "'except <cell> not in chain' = that cell is not used by the target.",
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
        "todas_formulas": todas,
        "falhas": len(falhas),
    }
    return texto, numeros


LIMITE_CONSTANTES = 200  # linhas de constantes por aba no --diff


def _igual(a, b):
    """Mesmo valor (ruido de ponto flutuante abaixo de 1e-9 nao conta)."""
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
    return fmt(a) == fmt(b)


def _abrir(fonte):
    """(wb, abas) de uma planilha, para o --diff."""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    wb = load_workbook(fonte)
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    wb_v = load_workbook(fonte, read_only=True, data_only=True)
    abas = ler_abas(wb, wb_v)[0]
    wb_v.close()
    return wb, {ws.title: (formulas, valores, cache) for ws, formulas, valores, cache in abas}


def _secao_diff(titulo, formulas, contexto, cache, falhas, aba):
    """Bloco compacto de um conjunto de formulas, com a conferencia de sempre."""
    if not formulas:
        return []
    linhas = secao_formulas(formulas, {}, "unchanged", contexto, cache)
    if not confere(linhas, esperado(formulas)):
        falhas.append("%s (%s)" % (aba, titulo))
        linhas = por_celula(formulas)
    return ["## " + titulo] + linhas


def converter_diff(antiga, nova):
    """Retorna (texto, numeros): so o que mudou da planilha antiga para a nova."""
    wb_a, abas_a = _abrir(antiga)
    wb_n, abas_n = _abrir(nova)
    falhas, corpo, resumo = [], [], []
    n_mudou = n_antes = n_const = n_iguais = 0
    vazio = ({}, {}, {})

    for nome in list(abas_n) + [t for t in abas_a if t not in abas_n]:
        if nome not in abas_a or nome not in abas_n:
            fs = (abas_n if nome in abas_n else abas_a)[nome][0]
            qtd = sum(1 for f in fs.values() if f[0] in "fa")
            resumo.append("- sheet '%s' %s (%d formulas)" % (nome, "added" if nome in abas_n else "removed", qtd))
            continue
        fa, va, ca = abas_a[nome]
        fn, vn, cn = abas_n[nome]
        ea, en = esperado(fa), esperado(fn)
        mudou = {pos: f for pos, f in fn.items() if f[0] in "fa" and ea.get(ref(*pos)) != en[ref(*pos)]}
        antes = {pos: f for pos, f in fa.items() if f[0] in "fa" and en.get(ref(*pos)) != ea[ref(*pos)]}
        const = []
        va = {q: x for q, x in va.items() if not (isinstance(x, str) and not x.strip())}  # "" = vazio
        vn = {q: x for q, x in vn.items() if not (isinstance(x, str) and not x.strip())}
        for pos in sorted(set(va) | set(vn)):
            a, n = va.get(pos), vn.get(pos)
            if pos in va and pos in vn:
                if not _igual(a, n):
                    const.append("%s %s -> %s" % (ref(*pos), fmt(a), fmt(n)))
            elif pos in vn:
                const.append("%s added %s" % (ref(*pos), fmt(n)))
            else:
                const.append("%s removed (was %s)" % (ref(*pos), fmt(a)))
        if not (mudou or antes or const):
            n_iguais += 1
            continue
        n_mudou += len(mudou)
        n_antes += len(antes)
        n_const += len(const)
        bloco = ["", "# Sheet '%s'" % nome]
        bloco += _secao_diff("Formulas new or changed", mudou, (fn, vn), cn, falhas, nome)
        bloco += _secao_diff("Formulas before (changed or removed cells)", antes, (fa, va), ca, falhas, nome)
        if const:
            extra = len(const) - LIMITE_CONSTANTES
            bloco += ["## Constants"] + const[:LIMITE_CONSTANTES] + (["... %d more" % extra] if extra > 0 else [])
        corpo += bloco

    nomes_a, nomes_n = set(nomes_definidos(wb_a)), set(nomes_definidos(wb_n))  # "NOME = destino"; destino novo = removido + novo
    nomes = ["removed: " + n for n in sorted(nomes_a - nomes_n)] + ["added: " + n for n in sorted(nomes_n - nomes_a)]

    cabeca = [
        "# excel_ai diff (old -> new)",
        "Only what changed. '## Formulas new or changed' is the NEW version, same format as the normal dump",
        "('<range> [label] (v value) <formula>', 'except' = cell inside the range that did not change).",
        "'## Formulas before' is the OLD formula of the cells that changed or disappeared.",
        "'A1 old -> new' is a constant that changed.",
        "Summary: %d formulas new or changed, %d old formulas replaced or removed, %d constants changed, %d sheet(s) without changes."
        % (n_mudou, n_antes, n_const, n_iguais),
    ]
    if falhas:
        cabeca.append("WARNING: compact view failed the self-check in: %s. Those parts list one formula per cell." % ", ".join(falhas))
    if resumo:
        cabeca += ["", "## Sheets"] + resumo
    if nomes:
        cabeca += ["", "## Defined names"] + nomes
    if not (corpo or resumo or nomes):
        cabeca.append("No changes found.")
    texto = "\n".join(cabeca + corpo) + "\n"
    return texto, {"novo": tokens(texto), "formulas": n_mudou, "antes": n_antes, "constantes": n_const,
                   "falhas": len(falhas)}


def main_diff(antiga, nova):
    for caminho in (antiga, nova):
        if not os.path.isfile(caminho):
            sys.stderr.write("erro: arquivo nao encontrado: %s\n" % ascii(caminho))
            return 2
    try:
        texto, n = converter_diff(antiga, nova)
    except Exception as e:
        sys.stderr.write("erro: nao consegui ler a planilha (%s: %s)\n" % (type(e).__name__, ascii(str(e))))
        return 2
    limpo = lambda c: re.sub(r"[^A-Za-z0-9]", "_", os.path.splitext(os.path.basename(c))[0])
    destino = os.path.join(os.path.dirname(os.path.abspath(nova)), "%s.vs_%s.diff.ai.txt" % (limpo(nova), limpo(antiga)))
    with io.open(destino, "w", encoding="utf-8", newline="\n") as arq:
        arq.write(texto)
    print("Tokens do diff:  %10s" % "{:,}".format(n["novo"]))
    print("Formulas novas/mudadas: %s | antigas trocadas/removidas: %s | constantes: %s" % (
        n["formulas"], n["antes"], n["constantes"]))
    if n["falhas"]:
        print("Conferencia das formulas: FALHOU em %d parte(s); elas sairam celula a celula." % n["falhas"])
    else:
        print("Conferencia das formulas: 100%%")
    print("Arquivo gerado: %s" % os.path.basename(destino).encode("ascii", "backslashreplace").decode())
    return 1 if n["falhas"] else 0


def main(caminho, alvo=None, macros_fonte=False):
    if not os.path.isfile(caminho):
        sys.stderr.write("erro: arquivo nao encontrado: %s\n" % ascii(caminho))
        return 2
    try:
        texto, n = converter(caminho, alvo, macros_fonte)
    except ValueError as e:
        sys.stderr.write("erro: %s\n" % ascii(str(e)))
        return 2
    except Exception as e:
        sys.stderr.write("erro: nao consegui ler a planilha (%s: %s)\n" % (type(e).__name__, ascii(str(e))))
        return 2
    sufixo = ""
    if alvo is not None:
        aba, cel = parte_aba(alvo.strip())
        sufixo = "." + re.sub(r"[^A-Za-z0-9]", "_", "%s_%s" % (aba or "ativa", cel.replace("$", "").upper()))
    destino = os.path.splitext(caminho)[0] + sufixo + ".ai.txt"
    with io.open(destino, "w", encoding="utf-8", newline="\n") as arq:
        arq.write(texto)
    print("Tokens hoje (celula a celula): %10s" % "{:,}".format(n["bruto"]))
    print("Tokens com excel_ai:           %10s" % "{:,}".format(n["novo"]))
    print("  avisos %s | mapa %s | formulas %s | dados %s" % tuple(
        "{:,}".format(n[k]) for k in ("avisos", "mapa", "formulas", "dados")))
    if n["novo"]:
        print("Reducao: %.1fx" % (n["bruto"] / float(n["novo"])))
    if alvo is not None:
        print("Cadeia: %s de %s formulas da planilha (%.1f%%)" % (
            "{:,}".format(n["total_formulas"]), "{:,}".format(n["todas_formulas"]),
            100.0 * n["total_formulas"] / max(n["todas_formulas"], 1)))
    if n["falhas"]:
        print("Conferencia das formulas: FALHOU em %d aba(s); essas abas sairam celula a celula." % n["falhas"])
    else:
        print("Conferencia das formulas: 100%% (%s formulas sem perda)" % "{:,}".format(n["total_formulas"]))
    print("Arquivo gerado: %s" % os.path.basename(destino).encode("ascii", "backslashreplace").decode())
    return 1 if n["falhas"] else 0


def main_pasta(pasta, macros_fonte=False):
    """Converte cada planilha da pasta com main(); um arquivo ruim nao para os outros."""
    nomes = sorted(n for n in os.listdir(pasta)
                   if n.lower().endswith((".xlsx", ".xlsm")) and not n.startswith("~$")
                   and os.path.isfile(os.path.join(pasta, n)))
    if not nomes:
        sys.stderr.write("erro: nenhuma planilha .xlsx/.xlsm na pasta: %s\n" % ascii(pasta))
        return 2
    codigos = []
    for n in nomes:
        print("== %s ==" % n.encode("ascii", "backslashreplace").decode())
        codigos.append(main(os.path.join(pasta, n), None, macros_fonte))
    print("Pasta: %d planilha(s), %d com problema" % (len(nomes), sum(1 for c in codigos if c)))
    return max(codigos)


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


SKILL_MD = r'''---
name: excel-ai
description: Read Excel workbooks (.xlsx/.xlsm) efficiently by converting them to compact, lossless text before analysing them. Use whenever the user asks about a spreadsheet, its formulas, calculation logic, cross-sheet references (XLOOKUP/PROCX), financial calculators (IRR/TIR, duration, price schedules, provisioning/PDD), why a cell has a value, what changed between two versions of a workbook, or wants a whole folder of workbooks read. Triggers - planilha, Excel, xlsx, xlsm, formula, PROCX, TIR, calculadora, diff de versoes.
license: MIT
---

# excel-ai: read spreadsheets as compact text

`scripts/excel_ai.py` turns a workbook into a short text file that keeps every formula and constant that matters and drops formatting noise. A 1.6M-token cell-by-cell dump of a large financial calculator becomes about 33k tokens (about 48x smaller). Formulas are verified lossless: the text is expanded back and compared with the original.

Never read an `.xlsx`/`.xlsm` by unzipping it or dumping cells with openpyxl yourself. Run the tool and read its output.

## Setup (once)

```
pip install openpyxl
```

Python 3.8+. One file, no network, nothing leaves the machine. The script is at `scripts/excel_ai.py` in this skill folder. Run it with that path (written below as `excel_ai.py`).

## Pick the mode

| The user wants... | Run |
|---|---|
| To understand or audit one workbook | `python excel_ai.py book.xlsx` |
| One number explained (final IRR, a total, a provision) | `python excel_ai.py book.xlsx --celula "Sheet!B5"` |
| What changed between two versions | `python excel_ai.py old.xlsx new.xlsx --diff` |
| Every workbook in a folder | `python excel_ai.py folder` |

Prefer `--celula` over the full dump whenever the question is about one result: it keeps only the cells that result depends on (for a real calculator, about 6% of the formulas and 2k tokens instead of 33k). Prefer `--diff` over reading two dumps.

Output files are written next to the input:

- full: `book.ai.txt`
- precedents: `book.<Sheet>_<CELL>.ai.txt` (non-alphanumerics become `_`)
- diff: `new.vs_old.diff.ai.txt`
- folder: one `.ai.txt` per workbook (not recursive; `~$*` lock files skipped)

After running, read the generated `.ai.txt` file (not the workbook). The console also prints tokens before/after and the self-check result.

## `--celula` details

- Reference forms: `'Sheet name'!B5`, `Sheet!B5`, `Sheet!$b$5`, or `B5` (active sheet). One cell only, never a range.
- It follows the chain through other sheets, ranges, whole columns/rows (limited to existing cells), defined names, table references (`Table[Col]`, `[@Col]`) and array formulas.
- Not followed, but listed in Warnings: `OFFSET`/`INDIRECT` results, external-file references, data-table (what-if) inputs, names that are not plain ranges, `#REF!`, and LET/LAMBDA variables. If Warnings says "some refs could not be followed", tell the user the chain may be incomplete and offer the full dump.
- Lines `except <cell> not in chain` are gaps inside a block: formulas that exist but are not used by the target.
- The header shows the target's cached value (what Excel last calculated).

## `--diff` details

- Sheets match by exact name. Added/removed sheets are listed with formula counts.
- `## Formulas new or changed` is the NEW version in the normal compact format; `## Formulas before` is the OLD formula of changed or removed cells. Constants appear as `A1 old -> new`, `A1 added X`, `A1 removed (was X)` (200 lines per sheet max). Defined names appear as `added:`/`removed:` (a changed destination is removed + added).
- A pure fill-down with the same relative formula is not a change. Float noise below 1e-9 and empty strings are ignored.
- Inserting rows or columns shows up as many changes (no move detection). Data-table cells are not compared. Say so when the diff looks huge.
- "No changes found." means formulas, constants and names are identical.

## Reading the output

Read `references/format.md` for the full legend. Short version:

- `## Warnings` first: `#REF!`, error values, OFFSET/INDIRECT, macros, uncalculated files. Mention relevant ones to the user.
- `## Map`: each sheet (dimensions, HIDDEN, formula count, which sheets it uses), tables, defined names.
- Per sheet `## Formulas`: `E2:E361 [Juros] (v 12.5) =C2*G2` means the formula of the first cell, copied down/across the range with references shifting. `[Juros]` is the nearest text header (a hint). `(v 12.5)` is the value Excel last calculated for the first cell (a hint, blocks of 2+ cells only). `  except E9 =...` overrides one cell. `same formulas as 'X', except:` means a copy of sheet X with listed differences.
- Per sheet `## Data`: constants, with long number runs summarized as `[100 numbers] min .. max .. first: ..`.
- Per sheet `## Formats` (only when present): cells Excel shows as `percent`, `date` or `currency`, as ranges (`percent: D2:D101, F2`). Use it to tell a rate from an amount. Only cells with a value are listed.
- Per sheet `## Notes` (only when present): the author's cell comments, `B5 note: <text>` (200 characters max, 100 per sheet). Treat them as the author's own explanation of assumptions.
- Per sheet `## Pivots` (only when present): each pivot table with where it sits, its source and its row, column, filter and value fields (`values: Soma de Valor = sum(Valor)`). The numbers it produced are ordinary cells in `## Data`; this section says they are a summary of the source. A warning appears when it refreshes on open.
- After the last sheet, `# Workbook` holds workbook-level sections (only when present). `## External links`: `[1] taxas.xlsx, sheets: T, 3 formulas` says `[1]` in formulas (`=[1]T!B2`) is that file, shown by name only (no folder), with the cached sheet names and how many formulas use it. The linked file is never opened. Warnings flag links without saved values and links to URLs.
- `## Connections` (in `# Workbook`, when present): one line per data connection: `'Name' (ODBC), source Server=db;Database=x, command SELECT ..., refreshes on open`. Sheets fed by a connection hold loaded data, not typed inputs; it may be stale. Only server/database/provider/driver/data-source parts of the connection string are kept, paths reduced to file names, URLs without query. Credentials never appear in the connection string; the `command` text is free text, redacted best-effort. Power Query connections (one per query) are counted in a single `Power Query: N connections` line; the queries themselves are in `## Queries`.
- `## Queries` (in `# Workbook`, when present): each Power Query by name, where it loads (`loads to: sheet table X`, `connection only`, `data model`) and its M code, indented, as written (2000 characters max, 50 queries). The code is not run or interpreted: read `Source = ...` to see where data comes from. Paths are reduced to file names, URLs lose their query, values after password/token keys are masked, best-effort. `Queries present but unreadable` means the file has Power Query this tool could not unpack.
- `## Macros` (in `# Workbook`, when the file has VBA): by default only `modules: A, B, C` (names). With `--macros-fonte`: one block per module `- Module1 (module): Sub A, Function B`, then `writes to cells:` (the distinct cell targets the code assigns, e.g. `Sheet1.Range("B2").Value`) and `source:` (indented, as written, 4000 characters per module and 20000 in total). The code is never run. A password-protected project lists names only. Use it to see which values a macro writes instead of a formula. Paths, URLs and values after password/token keys are masked, best-effort; review before sharing.

Reason about the logic from the formulas: say what each block computes, which sheets feed it, and where inputs live. Use cached values `(v ...)` only as sanity checks; they may be stale.

## Exit codes and problems

- `0` ok. `1` the lossless self-check failed on some sheet (that sheet is listed one formula per cell, still correct, just larger). `2` bad arguments, missing file, unreadable workbook, or a folder with no workbooks.
- "No calculated values saved": the file was never opened in Excel, so cached values and error checks are missing.
- "File has VBA macros": some values may come from a macro, which this tool does not read.
- `.xls` (old format) and `.xlsb` are not supported; ask the user to save as `.xlsx`.
- Large workbooks take 15-40 seconds. Do not interrupt.

## Rules

1. Run the tool before answering any question about a workbook's contents.
2. Never paste real workbook data into public places (issues, PRs, public repos). Workbooks are usually confidential.
3. Do not edit the workbook with this tool; it is read-only.
4. If the user's question is about one result, start with `--celula`. If it is about differences between versions, use `--diff`.
5. When the output says a part could not be followed or verified, say it plainly instead of guessing.
'''

FORMAT_MD = r'''# excel_ai output format

## File layout

```
# excel_ai dump
<how-to-read lines>

## Warnings
- ...

## Map
- 'Sheet' A1:J101, 303 formulas, uses 'Other' (200)
  table Tabela1 A1:B4 columns: Nome, Valor
Defined names:
  TAXA2024 = Premissas!$B$1

# Sheet 'Fluxo'
## Formulas
...
## Data
...
```

## Formula lines

| Line | Meaning |
|---|---|
| `D2:D101 =B2*C2` | Formula of D2; D3..D101 are the same formula filled down (relative references shift as in Excel). |
| `C2:E2 =C1*2` | Same, filled across. |
| `D2:D101 [Juros] =B2*C2` | `[Juros]` is the nearest text header: above for columns, left for rows, up to 2 empty cells away. A hint, not part of the formula. |
| `D2:D101 [Juros] (v 12.5) =B2*C2` | `(v 12.5)` is the cached value (last Excel result) of the FIRST cell. Only blocks of 2+ cells. Hint only. |
| `  except D30 =B30*C30*2` | D30 inside the block above has a different formula. |
| `  except D30 value 5` | D30 is a constant inside the block. |
| `  except D30 blank` | D30 is empty. In `--celula` mode: `not in chain` (exists but not used by the target). In a copy-sheet: `unchanged`. |
| `E1 array E1:E2 {=Z1:Z2*2}` | Array formula anchored at E1 covering E1:E2. Other cells of the range are not listed. |
| `same formulas as 'Fluxo', except:` | The sheet is a copy (80%+ equal) of that sheet; only differing cells follow. `A1 no formula` means the cell has no formula here. |

What-if data-table cells are not listed as lines; they only appear counted in Warnings.

Internal prefixes (`_xlfn.`) are removed: `_xlfn.XLOOKUP` is shown as `XLOOKUP` (PROCX in Portuguese Excel).

## Data lines

| Line | Meaning |
|---|---|
| `A1 Parcela \| B1 Saldo` | Constants on a row, `cell value` separated by `\|`. |
| `A2:A101 [100 numbers] min 1 max 100 first: 1, 2, 3` | Long run of numbers/dates summarized. Runs of up to 10 stay complete. |
| `A2:A69 [68 texts, 68 distinct] first: a, b, c` | Long run of texts (over 50) summarized. Shorter runs stay complete. |
| `A5:A9 x (x5)` | Same text repeated. |

## Formats and Notes

```
## Formats
percent: D2:D101, F2
currency: B2:B4
## Notes
A2 note: taxa anual em %
```

`percent` / `date` / `currency` come from the cell's number format in Excel; other formats (General, plain numbers, text) are not listed. Ranges run down a column; long lists end with `(+N more)`. `## Notes` holds cell comments verbatim (whitespace collapsed). With `--celula`, only the cells in the chain are listed. `--diff` ignores both.

## Pivots

```
## Pivots
- 'Piv1' at A3:B8, source 'Dados'!A1:C6
  rows: Mes
  columns: Values
  filters: Reg
  values: Soma de Valor = sum(Valor)
```

One block per pivot table, after `## Notes`. `source` is a sheet range, `table or name X` (Excel does not say which), or the cache type when it is not a worksheet. `Values` in `columns` or `rows` is the values axis. Only the definition is shown, never a re-computation. `--celula` and `--diff` ignore pivots.

## Workbook sections

```
# Workbook
## Connections
'Vendas' (ODBC), source Driver={SQL Server};Server=db;Database=x, command SELECT ..., refreshes on open
## Queries
- Vendas, loads to: sheet table Vendas
    let
        Source = Excel.Workbook(File.Contents("dados.xlsx"))
    in
        Source
## Macros
modules: Module1, Sheet1
(source not shown: run with --macros-fonte)
## External links
[1] taxas.xlsx, sheets: T, 3 formulas
```

Emitted once, after the last sheet, in this order when present: Connections, Queries, External links, Macros. `[N]` is the 1-based position of the link; formulas keep `[N]` unchanged. Paths are reduced to file names and URL query strings are dropped. `--celula` and `--diff` omit this block.

## Warnings you may see

- `N formulas contain #REF!` broken reference.
- `N formulas show an error value (#N/A, #DIV/0!...)` from cached values.
- `N formulas use OFFSET/INDIRECT` target decided at run time; not followable statically.
- `N data-table (what-if) cells`.
- `File has VBA macros`.
- `VBA project is password protected: source never shown`.
- `connection 'X' refreshes when the file opens`.
- `Queries present but unreadable` (Power Query exists, could not be unpacked).
- `external link [N] has no saved values` / `points to a URL`.
- `pivot 'X' refreshes when the file opens` its numbers may differ from the saved ones.
- `No calculated values saved in the file`.
- `compact view failed the self-check on N sheet(s)` those sheets are one formula per cell.
- `--celula` only: `unresolved` / `formula not parsed` refs that could not be followed.

## Diff file

```
# excel_ai diff (old -> new)
Summary: N formulas new or changed, M old formulas replaced or removed, K constants changed, S sheet(s) without changes.
## Sheets          (added/removed sheets)
## Defined names   (added:/removed:)
# Sheet 'X'
## Formulas new or changed
## Formulas before (changed or removed cells)
## Constants
```
'''


def instalar_skill(destino=None, forcar=False):
    """Grava a skill do Kiro (SKILL.md, references/format.md e uma copia deste script). Pula se ja existe."""
    destino = destino or os.path.expanduser("~/.kiro/skills/excel-ai")
    if os.path.exists(destino) and not forcar:
        print("Skill ja instalada em %s (use --forcar para sobrescrever)" % destino)
        return 0
    for rel, texto in (("SKILL.md", SKILL_MD), ("references/format.md", FORMAT_MD)):
        caminho = os.path.join(destino, rel)
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(texto)
    os.makedirs(os.path.join(destino, "scripts"), exist_ok=True)
    shutil.copyfile(os.path.abspath(__file__), os.path.join(destino, "scripts", "excel_ai.py"))
    print("Skill instalada em %s" % destino)
    return 0


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
    assert "D2:D79 [Juros] =B2*C2" in fluxo, fluxo
    assert "  except D50 =B50*0" in fluxo
    assert "  except D60 value 7" in fluxo
    assert "  except D70 blank" in fluxo
    assert "D80 =1" in fluxo and "D81 =2" in fluxo and "D82:D101 =B82*C82" in fluxo  # formula acima: sem rotulo
    assert "C2:C101 [Taxa] =XLOOKUP(" in fluxo and "_xlfn" not in fluxo
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

    # rotulo: cabecalho acima (coluna) ou a esquerda (linha); pula ate 2 vazias
    fs = {(r, 2): ("f", "=A%d" % r, relativa("=A%d" % r, r, 2)) for r in range(5, 8)}
    assert blocos(fs, {(2, 2): "Juros"}, True)[0] == "B5:B7 [Juros] =A5"
    assert blocos(fs, {(1, 2): "Juros"}, True)[0] == "B5:B7 =A5"  # 3 vazias: longe demais
    assert blocos(fs, {(4, 2): 12}, True)[0] == "B5:B7 =A5"  # numero nao e rotulo
    assert blocos(fs, {(4, 2): "x [a]  muito   comprido " * 3}, True)[0] == "B5:B7 [x a muito comprido x a muito c] =A5"
    fs[(4, 2)] = ("f", "=1", "=1")
    assert blocos(fs, {(3, 2): "Juros"}, True)[1] == "B5:B7 =A5"  # formula acima para a busca
    fh = {(3, c): ("f", "=%s1" % get_column_letter(c), relativa("=%s1" % get_column_letter(c), 3, c)) for c in range(3, 6)}
    assert blocos(fh, {(3, 1): "Total"}, False)[0] == "C3:E3 [Total] =C1"
    assert blocos(fh, {(3, 2): "%"}, False)[0] == "C3:E3 =C1"  # unidade nao e rotulo
    assert blocos(fh, {(3, 2): "x"}, False)[0] == "C3:E3 =C1"  # uma letra so nao e rotulo
    assert blocos(fh, {(3, 2): "#N/A"}, False)[0] == "C3:E3 =C1"  # erro nao e rotulo
    assert blocos(fh, {(3, 2): "  ", (3, 1): "Total"}, False)[0] == "C3:E3 [Total] =C1"  # espaco = vazio
    um = {(4, 3): ("f", "=A1", relativa("=A1", 4, 3))}
    assert blocos(um, {(4, 2): 5, (3, 3): "TIR"}, True) == ["C4 =A1"]  # numero a esquerda: sem rotulo
    assert blocos(um, {(3, 3): "TIR"}, True) == ["C4 [TIR] =A1"]  # nada a esquerda: olha acima
    assert blocos(um, {(4, 2): "Taxa", (3, 3): "TIR"}, True) == ["C4 [Taxa] =A1"]  # esquerda ganha
    assert expandir(["B5:B7 [Juros] =A5"]) == expandir(["B5:B7 =A5"])

    # valor calculado: so o 1o da faixa, depois do rotulo; sem cache ou None = sem dica
    fs = {(r, 2): ("f", "=A%d" % r, relativa("=A%d" % r, r, 2)) for r in range(5, 8)}
    assert blocos(fs, {(4, 2): "Juros"}, True, cache={(5, 2): 12.5, (6, 2): 99})[0] == "B5:B7 [Juros] (v 12.5) =A5"
    assert blocos(fs, {}, True, cache={(5, 2): 3})[0] == "B5:B7 (v 3) =A5"
    assert blocos({(5, 2): fs[(5, 2)]}, {}, True, cache={(5, 2): 3})[0] == "B5 =A5"  # celula sozinha: sem dica
    assert blocos(fs, {}, True, cache={(6, 2): 3})[0] == "B5:B7 =A5"  # 1a celula sem valor
    assert blocos(fs, {}, True, cache={(5, 2): None})[0] == "B5:B7 =A5"
    assert blocos(fs, {}, True, cache={(5, 2): "a)b\nc  " + "x" * 30})[0] == "B5:B7 (v ab c xxxxxxxxxxxxxxx) =A5"
    assert blocos(fs, {}, True, cache={(5, 2): "#N/A"})[0] == "B5:B7 (v #N/A) =A5"
    assert blocos(fs, {}, True, cache={(5, 2): True})[0] == "B5:B7 (v TRUE) =A5"
    assert blocos(fs, {}, True)[0] == "B5:B7 =A5"
    fx = dict(fs)
    fx[(6, 2)] = ("f", "=Z1", relativa("=Z1", 6, 2))
    assert not any("(v" in l for l in blocos(fx, {}, True, cache={(5, 2): 1, (6, 2): 2}) if l.startswith("  except"))
    assert expandir(["B5:B7 [Juros] (v 12.5) =A5"]) == expandir(["B5:B7 =A5"])
    assert expandir(["B5 (v 1) =A5"]) == {"B5": "=A5"}
    assert not confere(["B5:B7 (v 1) =A5"], {"B5": "=A5", "B6": "=A6", "B7": "=A8"})
    assert not confere(["B5:B7 (v 1 =A5"], {"B5": "=A5", "B6": "=A6", "B7": "=A7"})
    certo = {"D%d" % r: "=B%d*C%d" % (r, r) for r in range(2, 11)}
    assert confere(["D2:D10 [J] =B2*C2"], certo) and not confere(["D2:D10 [J] =B2*C3"], certo)
    assert not confere(["D2:D10 [J =B2*C2"], certo)  # rotulo mal fechado nao passa

    # conferencia pega uma linha estragada
    linhas = ["D2:D10 =B2*C2"]
    certo = {"D%d" % r: "=B%d*C%d" % (r, r) for r in range(2, 11)}
    assert expandir(linhas) == certo
    assert expandir(["D2:D10 =B2*C3"]) != certo
    assert not confere(["linha que nao e formula"], certo)

    # formula com quebra de linha (Alt+Enter) sai numa linha so e confere
    wb = Workbook()
    wb.active["A1"], wb.active["B1"] = 1, "=IF(A1>0,\n1,2)"
    buf = io.BytesIO()
    wb.save(buf)
    texto, n = converter(buf)
    assert "B1 =IF(A1>0, 1,2)" in texto and n["falhas"] == 0

    # aba copiada com matriz de faixa diferente nao passa como igual
    from openpyxl.worksheet.formula import ArrayFormula
    wb = Workbook()
    for nome, faixa in (("S1", "C1:C3"), ("S2", "C1:C2")):
        ws = wb.create_sheet(nome)
        for r in range(1, 30):
            ws["A%d" % r] = "=B%d" % r
        ws["C1"] = ArrayFormula(faixa, "=A1:A3*2")
    buf = io.BytesIO()
    wb.save(buf)
    texto, n = converter(buf)
    assert "C1 array C1:C2 {=A1:A3*2}" in texto.split("# Sheet 'S2'")[1] and n["falhas"] == 0

    # plano B: se a conferencia falha, a aba sai celula a celula, com aviso, e main() sai com 1
    import contextlib
    import tempfile
    pasta = tempfile.mkdtemp()
    caminho = os.path.join(pasta, "fix.xlsx")
    with open(caminho, "wb") as arq:
        arq.write(_fixture().getvalue())
    with contextlib.redirect_stdout(io.StringIO()):
        assert main(caminho) == 0 and os.path.isfile(os.path.join(pasta, "fix.ai.txt"))
        original = globals()["expandir"]
        globals()["expandir"] = lambda linhas, base=None: {}
        try:
            texto, n = converter(caminho)
            assert main(caminho) == 1
        finally:
            globals()["expandir"] = original
    # 3 abas com formulas (Fluxo, Meses, Fluxo B); abas sem formula nao tem o que conferir
    assert n["falhas"] == 3 and "WARNING: compact view failed the self-check on 3 sheet(s)" in texto
    assert "D3 =B3*C3" in texto  # uma linha por formula
    with contextlib.redirect_stderr(io.StringIO()):
        assert main(os.path.join(pasta, "nao_existe.xlsx")) == 2
    # Translator do openpyxl apagava ABA!#REF!; tradutor() protege
    assert tradutor("=XLOOKUP(A2,Base!#REF!,'P 1'!#REF!)+#REF!", "B2")("B3") == "=XLOOKUP(A3,Base!#REF!,'P 1'!#REF!)+#REF!"

    # cadeia de precedentes (--celula): so o que a celula usa, entre abas
    texto, n = converter(_fixture(), "Fluxo!D5")
    assert "Precedents of 'Fluxo'!D5" in texto, texto[:400]
    fl = texto.split("# Sheet 'Fluxo'")[1].split("# Sheet ")[0]
    assert "D5 =B5*C5" in fl and "B5 =Premissas!$B$2-A5" in fl and "C5 =XLOOKUP(" in fl, fl
    assert "D6" not in fl and "Meses" not in texto and "Fluxo B" not in texto
    assert "A3 Carencia | B3 6" in texto  # coluna inteira do XLOOKUP entra
    assert n["falhas"] == 0 and n["total_formulas"] == 3
    texto, n = converter(_fixture(), "fluxo!$j$2")  # nome definido + outra aba
    assert "TAXA2024 = Premissas!$B$1" in texto and "A1 Nome" in texto and "B1 0.012" in texto
    texto, n = converter(_fixture(), "Fluxo!I2")  # tabela: so a coluna Valor
    extra = texto.split("# Sheet 'Plan Extra 2'")[1]
    assert "B2" in extra and "item0" not in extra, extra
    for ruim in ("Nada!A1", "Fluxo!A1:B2", "Fluxo!ZZZZ1"):
        try:
            converter(_fixture(), ruim)
            assert False, ruim
        except ValueError:
            pass
    from openpyxl.worksheet.table import Table
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "Valor"
    for r in range(2, 13):
        ws["A%d" % r], ws["B%d" % r] = "=B%d*2" % r, r
    w2 = wb.create_sheet("T2")
    w2["D1"], w2["E1"] = "=E1", "=D1"  # ciclo
    w2["F1"], w2["G1"] = "=Nope!A1", '=INDIRECT("A1")'
    w2["H1"] = ArrayFormula("H1:H3", "=B2:B4*2")
    w2["I1"] = "=H2"  # meio da matriz puxa a ancora
    w2["K1"], w2["L1"], w2["K2"], w2["K3"] = "Qtd", "Dobro", 7, 8
    w2["L2"], w2["L3"] = "=[@Qtd]*2", "=[@Qtd]*2"
    w2.add_table(Table(displayName="T", ref="K1:L3"))
    from openpyxl.workbook.defined_name import DefinedName
    try:
        wb.defined_names.add(DefinedName("LOOP", attr_text="LOOP"))  # nome que aponta para si
    except AttributeError:
        wb.defined_names.append(DefinedName("LOOP", attr_text="LOOP"))
    ws["C1"] = "=SUM(A2:A3)+SUM(A5:A12)+T2!D1+T2!F1+T2!G1+T2!I1+T2!L2+LOOP"
    buf = io.BytesIO()
    wb.save(buf)
    texto, n = converter(buf, "C1")
    assert "A2:A12 [Valor] =B2*2" in texto and "  except A4 not in chain" in texto, texto
    assert "D1 =E1" in texto and "E1 =D1" in texto and "H1 array H1:H3" in texto
    assert "- 'T2': unresolved reference Nope!A1 in F1" in texto and "unresolved reference LOOP" in texto and "OFFSET/INDIRECT" in texto
    assert "K2 7" in texto and "L3 =" not in texto and n["falhas"] == 0
    # nome curto nao e coluna; nome de aba ganha do global; constante nao avisa; B1# avisa; meio de matriz
    wb = Workbook()
    s1, s2 = wb.active, wb.create_sheet("S2")
    s1.title = "S1"
    s1["Z1"], s2["A1"], s2["B1"] = 5, 1, 100
    for nome, texto in (("TIR", "S1!$Z$1"), ("Rate", "S2!$B$1"), ("Taxa", "0.05")):
        try:
            wb.defined_names.add(DefinedName(nome, attr_text=texto))
        except AttributeError:
            wb.defined_names.append(DefinedName(nome, attr_text=texto))
    try:
        s1.defined_names.add(DefinedName("Rate", attr_text="S2!$A$1"))
    except AttributeError:
        s1.defined_names["Rate"] = DefinedName("Rate", attr_text="S2!$A$1")
    s1["A1"] = "=TIR*Rate*Taxa+C1"
    s1["C1"] = "=SUM(D1#)"
    s1["E1"] = ArrayFormula("E1:E2", "=Z1:Z2*2")
    buf = io.BytesIO()
    wb.save(buf)
    texto, n = converter(buf, "S1!A1")
    assert "Z1 5" in texto and "A1 1" in texto and "B1 100" not in texto, texto
    assert "unresolved" not in texto and "formula not parsed" in texto and "see Warnings" in texto
    texto, n = converter(buf, "S1!E2")
    assert "E1 array E1:E2" in texto
    import subprocess
    uso = subprocess.run([sys.executable, os.path.abspath(__file__), caminho, "--celula", "Fluxo!D5"],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert uso.returncode == 0 and b"Cadeia: 3 de 618 formulas" in uso.stdout, uso.stdout
    uso = subprocess.run([sys.executable, os.path.abspath(__file__), caminho, "x", "y"],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert uso.returncode == 2 and b"uso:" in uso.stderr
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        assert main(caminho, "'Fluxo'!D5") == 0
        assert os.path.isfile(os.path.join(pasta, "fix.Fluxo_D5.ai.txt"))
        assert main(caminho, "Nada!A1") == 2

    # --diff: duas versoes
    from openpyxl.workbook.defined_name import DefinedName as DN

    def livro(edita=None):
        w = Workbook()
        a = w.active
        a.title = "S"
        a["A1"], a["B1"] = "Item", "Valor"
        for r in range(2, 12):
            a["A%d" % r] = "x%d" % r
            a["B%d" % r] = r
            a["C%d" % r] = "=B%d*2" % r
        a["D1"] = "=SUM(C2:C11)"
        w.create_sheet("T")["A1"] = "=S!D1"
        w.defined_names["Taxa"] = DN("Taxa", attr_text="S!$B$2")
        if edita:
            edita(w)
        b = io.BytesIO()
        w.save(b)
        return b

    velho = livro()
    texto, n = converter_diff(velho, livro())
    assert "No changes found." in texto and n["formulas"] == n["antes"] == n["constantes"] == 0 and n["falhas"] == 0, texto

    def v2(w):
        a = w["S"]
        for r in range(5, 9):
            a["C%d" % r] = "=B%d*3" % r  # bloco de 4 mudou
        a["C11"] = 7  # formula virou constante
        a["B2"] = 99  # constante mudou
        a["E1"] = "novo"  # constante nova
        a["A3"] = None  # constante removida
        w["T"]["A1"] = "=S!D1+1"
        w.create_sheet("U")["A1"] = "=1+1"
        del w["S"].parent.defined_names["Taxa"]
        w.defined_names["Outra"] = DN("Outra", attr_text="S!$B$3")

    texto, n = converter_diff(velho, livro(v2))
    assert n["falhas"] == 0, texto
    assert "C5:C8 =B5*3" in texto and "C5:C8 =B5*2" in texto, texto  # novo e antigo
    assert "C11 =B11*2" in texto and "C11 7" not in texto.split("## Constants")[0]
    assert "B2 2 -> 99" in texto and "E1 added novo" in texto and "A3 removed (was x3)" in texto and "C11 added 7" in texto, texto
    assert "A1 =S!D1+1" in texto and "A1 =S!D1\n" in texto
    assert "- sheet 'U' added (1 formulas)" in texto
    assert "removed: Taxa = S!$B$2" in texto and "added: Outra = S!$B$3" in texto, texto
    assert "C2:C4" not in texto and "SUM" not in texto  # o que nao mudou fica de fora
    assert n["formulas"] == 5 and n["antes"] == 6 and n["constantes"] == 4, n
    # preencher para baixo igual nao e mudanca (mesma formula relativa)
    texto, n = converter_diff(velho, livro(lambda w: w["S"].__setitem__("C2", "=B2*2")))
    assert "No changes found." in texto
    # ruido: "" vira vazio, float em 1e-12 nao conta
    def ruido(w):
        w["S"]["F1"] = ""
        w["S"]["B3"] = 3.0000000000001
    texto, n = converter_diff(velho, livro(ruido))
    assert "No changes found." in texto, texto
    # mais de 200 constantes: corta
    def muitas(w):
        for r in range(1, 301):
            w["T"]["B%d" % r] = r
    texto, n = converter_diff(velho, livro(muitas))
    assert "... 100 more" in texto and texto.count("added") == 200 and n["constantes"] == 300
    # conferencia que falha cai para uma linha por celula
    guarda = confere
    globals()["confere"] = lambda linhas, meu, base=None: False
    try:
        texto, n = converter_diff(velho, livro(v2))
    finally:
        globals()["confere"] = guarda
    assert n["falhas"] > 0 and "WARNING: compact view failed" in texto and "C5 =B5*3" in texto, texto
    # CLI
    pa, pn = os.path.join(pasta, "old v1.xlsx"), os.path.join(pasta, "new v2.xlsx")
    for caminho_x, b in ((pa, velho), (pn, livro(v2))):
        with open(caminho_x, "wb") as arq:
            arq.write(b.getvalue())
    uso = subprocess.run([sys.executable, os.path.abspath(__file__), pa, pn, "--diff"],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert uso.returncode == 0 and b"100%" in uso.stdout, uso.stdout + uso.stderr
    assert os.path.isfile(os.path.join(pasta, "new_v2.vs_old_v1.diff.ai.txt"))
    uso = subprocess.run([sys.executable, os.path.abspath(__file__), pa, os.path.join(pasta, "nada.xlsx"), "--diff"],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert uso.returncode == 2
    uso = subprocess.run([sys.executable, os.path.abspath(__file__), pa, pn, "--celula"],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert uso.returncode == 2 and b"uso:" in uso.stderr

    # pasta inteira
    sub = os.path.join(pasta, "lote")
    os.makedirs(os.path.join(sub, "fundo"))
    for nome_x in ("b.xlsx", "a.XLSM", os.path.join("fundo", "c.xlsx")):
        with open(os.path.join(sub, nome_x), "wb") as arq:
            arq.write(_fixture().getvalue())
    for nome_x in ("~$a.xlsx", "ruim.xlsx", "nota.txt"):
        with open(os.path.join(sub, nome_x), "wb") as arq:
            arq.write(b"nao e planilha")
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
        assert main_pasta(sub) == 2
    assert os.path.isfile(os.path.join(sub, "a.ai.txt")) and os.path.isfile(os.path.join(sub, "b.ai.txt"))
    assert not os.path.exists(os.path.join(sub, "ruim.ai.txt")) and not os.path.exists(os.path.join(sub, "fundo", "c.ai.txt"))
    assert "Pasta: 3 planilha(s), 1 com problema" in out.getvalue() and "== a.XLSM ==" in out.getvalue(), out.getvalue()
    vazia = os.path.join(pasta, "vazia")
    os.makedirs(vazia)
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        assert main_pasta(vazia) == 2
    uso = subprocess.run([sys.executable, os.path.abspath(__file__), os.path.join(sub, "fundo")],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert uso.returncode == 0 and b"Pasta: 1 planilha(s), 0 com problema" in uso.stdout, uso.stdout

    # macro: zip com vbaProject.bin
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as arq:
        arq.writestr("xl/vbaProject.bin", b"x")
    assert tem_vba(z) and not tem_vba(_fixture())
    # skill do Kiro: instala, pula se ja existe, --forcar sobrescreve
    destino = os.path.join(pasta, "skill")
    with contextlib.redirect_stdout(io.StringIO()):
        assert instalar_skill(destino) == 0
    for rel in ("SKILL.md", "references/format.md", "scripts/excel_ai.py"):
        assert os.path.isfile(os.path.join(destino, rel)), rel
    with open(os.path.join(destino, "SKILL.md"), "w") as f:
        f.write("meu")
    with contextlib.redirect_stdout(io.StringIO()):
        instalar_skill(destino)
    assert open(os.path.join(destino, "SKILL.md")).read() == "meu"
    with contextlib.redirect_stdout(io.StringIO()):
        instalar_skill(destino, forcar=True)
    assert open(os.path.join(destino, "SKILL.md")).read() == SKILL_MD
    # formatos e notas: fatos do arquivo; planilha sem nenhum dos dois nao ganha secao
    from openpyxl.comments import Comment
    wbx = Workbook()
    wsx = wbx.active
    wsx.title = "Taxas"
    for r in range(2, 5):
        wsx.cell(r, 1, 0.1 * r).number_format = "0.00%"
        wsx.cell(r, 2, 100 * r).number_format = '"R$" #,##0.00'
    wsx["C2"] = datetime.date(2024, 1, 31)
    wsx["C2"].number_format = "dd/mm/yyyy"
    wsx["A2"].comment = Comment("taxa   anual\nem %", "x")
    wsx["D1"] = 5
    wsy = wbx.create_sheet("Limpa")
    wsy["A1"] = 7
    bio = io.BytesIO()
    wbx.save(bio)
    txt, _ = converter(bio)
    taxas = txt.split("# Sheet 'Taxas'")[1].split("# Sheet 'Limpa'")[0]
    assert "## Formats\ncurrency: B2:B4\ndate: C2\npercent: A2:A4" in taxas, taxas
    assert "## Notes\nA2 note: taxa anual em %" in taxas, taxas
    assert "## Formats" not in txt.split("# Sheet 'Limpa'")[1] and "## Notes" not in txt.split("# Sheet 'Limpa'")[1]
    assert texto_nota(Comment("[Comentário encadeado]\n\nAviso. Saiba mais: http://x\n\nComentário:\n    Fraude = 0,10%", "x")) == "Fraude = 0,10%"
    assert tipo_formato("General") == "" and tipo_formato("0.00") == "" and tipo_formato("[$-409]m/d/yyyy") == "date"
    # pivots: leitura de objetos em memoria + ida e volta pelo arquivo; sem pivot nada muda
    from openpyxl.pivot.cache import CacheDefinition, CacheField, CacheSource, WorksheetSource
    from openpyxl.pivot.table import DataField, Location, PageField, RowColField, TableDefinition
    wbp = Workbook()
    dp = wbp.active
    dp.title = "Dados"
    dp.append(["Mes", "Valor", "Reg"])
    for i in range(5):
        dp.append(["m%d" % i, i, "a"])
    rp = wbp.create_sheet("Resumo")
    cachep = CacheDefinition(cacheSource=CacheSource(type="worksheet", worksheetSource=WorksheetSource(ref="A1:C6", sheet="Dados")),
                             cacheFields=[CacheField(name="Mes"), CacheField(name="Valor"), CacheField(name="Reg")], refreshOnLoad=True)
    pivp = TableDefinition(name="Piv1", cacheId=0, dataCaption="Values", location=Location(ref="A3:B8", firstHeaderRow=1, firstDataRow=1, firstDataCol=1),
                           rowFields=[RowColField(x=0)], colFields=[RowColField(x=-2)], pageFields=[PageField(fld=2)],
                           dataFields=[DataField(name="Soma de Valor", fld=1, subtotal="sum"), DataField(name="N", fld=0, subtotal="count")])
    pivp.cache = cachep
    rp.add_pivot(pivp)
    bio = io.BytesIO()
    wbp.save(bio)
    txt, _ = converter(bio)
    resumo = txt.split("# Sheet 'Resumo'")[1]
    assert ("## Pivots\n- 'Piv1' at A3:B8, source 'Dados'!A1:C6\n  rows: Mes\n  columns: Values\n  filters: Reg\n"
            "  values: Soma de Valor = sum(Valor), N = count(Mes)") in resumo, resumo
    assert "- 'Resumo': pivot 'Piv1' refreshes when the file opens" in txt
    assert "'## Pivots' lists pivot tables" in txt and "## Pivots" not in txt.split("# Sheet 'Dados'")[1].split("# Sheet 'Resumo'")[0]
    assert "## Pivots" not in converter(bio, "Dados!A1")[0]  # --celula: nada de pivot
    assert "## Pivots" not in converter(_fixture())[0] and "# Workbook" not in converter(_fixture())[0]
    # formatador com objetos parciais: fonte por nome de tabela, so valores, indice fora do intervalo, sem cache
    pivp.cache = None
    assert secao_pivots(rp)[0][0].endswith("source unknown") and "no readable cache" in secao_pivots(rp)[1][0]
    c2 = CacheDefinition(cacheSource=CacheSource(type="worksheet", worksheetSource=WorksheetSource(name="Tabela1")), cacheFields=[CacheField(name="X")])
    p2 = TableDefinition(name="P2", cacheId=0, dataCaption="Values", location=Location(ref="A1:A2", firstHeaderRow=1, firstDataRow=1, firstDataCol=1),
                         colFields=[RowColField(x=9)], dataFields=[DataField(name="T", fld=0, subtotal="max")])
    p2.cache = c2
    l2, a2 = _pivot(p2)
    assert l2 == ["- 'P2' at A1:A2, source table or name Tabela1", "  columns: field 9", "  values: T = max(X)"] and not a2, l2
    p2.cache.cacheSource = CacheSource(type="external", connectionId=1)
    assert _pivot(p2)[0][0].endswith("source external (not a worksheet range)")
    rp._pivots = [object()]  # objeto ilegivel: aviso, sem excecao
    assert secao_pivots(rp) == (["- pivot unreadable"], ["- 'Resumo': a pivot could not be read"])

    # ajudantes compartilhados: planta segredos de verdade e confere que sumiram E que o resto ficou
    r = redigir("C:\\Users\\ana\\Docs\\taxas.xlsx")
    assert r == "taxas.xlsx", r
    assert redigir("\\\\srv\\share\\ana\\x.xlsx") == "x.xlsx" and redigir("/home/ana/x.xlsx") == "x.xlsx"
    assert redigir("file:///x/ana/y%20z.xlsx") == "y z.xlsx"
    r = redigir("https://user:pw9@h.com/p/a.xlsx?api_key=zzz9#frag9")
    assert r == "https://h.com/p/a.xlsx", r
    r = redigir('x [Password = "hunter2"] Server=db;PWD=hunter3; mail ana@corp.com; ApiKey=abc9')
    assert not any(k in r for k in ("hunter2", "hunter3", "abc9", "ana@")) and "Server=db" in r and "***@***" in r, r
    r = redigir('Web.Contents("https://h/p", [Headers=[Authorization="Bearer abc9"]])')
    assert "abc9" not in r and "https://h/p" in r, r
    assert redigir("ftp://u:pw9@h/x") == "ftp://h/x" and redigir("D:/Users/ana/x.xlsx") == "x.xlsx"
    assert "dXNl" not in redigir("Authorization: Basic dXNlcjpwdw==") and redigir("see http://[x").startswith("see http://")
    assert redigir("PWD=pa ss;Server=db") == "PWD=***;Server=db"
    import time
    t0 = time.time()
    for lixo in ("a" * 200000, "\\" * 3000000, "/" * 3000000, "a:" * 1000000):
        redigir(lixo)
    assert time.time() - t0 < 5, "redigir quadratico"
    assert redigir("//u9:p9@h/x") == "//h/x" and redigir("Data Source=u9:p9@host") == "Data Source=host"
    assert redigir("a@corp.com") == "***@***"
    assert redigir("=SUM(A1:A3)") == "=SUM(A1:A3)"
    z = _injetar(_fixture(), {"xl/extra.bin": b"zz"})
    assert zipfile.ZipFile(z).read("xl/extra.bin") == b"zz"
    load_workbook(z)
    assert ler_membro(zipfile.ZipFile(z), "xl/extra.bin") == b"zz"
    try:
        ler_membro(zipfile.ZipFile(z), "xl/extra.bin", 1)
        assert False, "teto nao aplicado"
    except ValueError:
        pass
    for ruim in (b'<?xml version="1.0"?><!DOCTYPE a [<!ENTITY e "x">]><a>&e;</a>', '<!DOCTYPE a><a/>'.encode("utf-16")):
        try:
            xml_seguro(ruim)
            assert False, "DOCTYPE aceito"
        except ValueError:
            pass
    assert xml_seguro(b"<a><b/></a>").tag == "a"
    assert bloco_livro({}) == [] and bloco_livro({"Macros": ["## M"], "Connections": ["## C"]}) == ["", "# Workbook", "## C", "## M"]
    # links externos: [N] -> arquivo, so o nome, contagem por indice, avisos
    from openpyxl.packaging.relationship import Relationship
    from openpyxl.workbook.external_link.external import ExternalBook, ExternalLink, ExternalSheetNames

    def _link(alvo, abas=("T",)):
        k = ExternalLink(externalBook=ExternalBook(sheetNames=ExternalSheetNames(list(abas)), id="rId1"))
        k.file_link = Relationship(type="externalLinkPath", Target=alvo, TargetMode="External", Id="rId1")
        return k
    wbl = Workbook()
    wl = wbl.active
    wl.title = "A"
    wl["A1"], wl["A2"], wl["A3"] = "=[1]T!B2", "=[1]T!B3+[1]T!B4", "=SUM([2]T!B2:B3)"
    wl["A4"] = "=Tab[Col]+{1,2}"
    wl["A5"] = "=[2]2020!A1"
    wl2 = wbl.create_sheet("B")
    wl2["A1"] = "=[1]T!B9"
    wbl._external_links += [_link("file:///C:\\Users\\ana\\Docs\\taxas%20x.xlsx"), _link("https://user:pw9@h.com/d/p.xlsx?sig=zzz9", ("T", "U")),
                            _link("..\\ana\\rel.xlsx?sig=zzz8")]
    bio = io.BytesIO()
    wbl.save(bio)
    txt, _ = converter(bio)
    wl_txt = txt.split("\n# Workbook\n")[1]
    assert wl_txt == ("## External links\n[1] taxas x.xlsx, sheets: T, 3 formulas\n"
                      "[2] https://h.com/d/p.xlsx, sheets: T, U, 2 formulas\n[3] rel.xlsx, sheets: T, 0 formulas\n"), wl_txt
    assert not any(x in txt for x in ("ana", "pw9", "zzz9", "zzz8", "Users"))
    assert "- external link [1] has no saved values" in txt and "- external link [2] points to a URL: https://h.com/d/p.xlsx" in txt
    assert txt.index("# Workbook") > txt.index("# Sheet 'B'") and "'## External links' (after the last sheet)" in txt
    assert "# Workbook" not in converter(bio, "A!A1")[0]
    wbl._external_links[0].file_link = None  # destino ilegivel: sem excecao
    assert links_externos(wbl, [])[0][0].startswith("[1] unknown target")
    # conexoes: lista permitida da string de conexao; segredos plantados nao saem, o resto fica
    cx = (b'<?xml version="1.0"?><connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          b'<connection id="1" name="Vendas" type="1" refreshOnLoad="1" background="1"><dbPr connection="Driver={SQL Server};Server=db;Database=x;UID=ana9;PWD=hunter2;ApiKey=abc9"'
          b' command="SELECT a FROM t -- PWD=hunter4"/></connection>'
          b'<connection id="2" name="Acc" type="5"><dbPr connection="Provider=Microsoft.ACE.OLEDB;Data Source=C:\\Users\\ana\\x.accdb;Extended Properties=hunter5"/></connection>'
          b'<connection id="3" name="Site" type="4"><webPr url="https://h.com/p?token=zzz9"/></connection>'
          b'<connection id="4" name="Arq" type="4"><textPr sourceFile="C:\\Users\\ana\\data.csv"/></connection>'
          b'<connection id="5" name="Q1" type="5"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Q1"/></connection></connections>')
    txt, _ = converter(_injetar(_fixture(), {"xl/connections.xml": cx}))
    cons = txt.split("\n# Workbook\n")[1]
    assert cons == ("## Connections\n"
                    "- 'Vendas' (ODBC), source Driver={SQL Server};Server=db;Database=x, command SELECT a FROM t -- PWD=***, refreshes on open, background refresh\n"
                    "- 'Acc' (OLE DB), source Provider=Microsoft.ACE.OLEDB;Data Source=x.accdb\n"
                    "- 'Site' (text), source https://h.com/p\n- 'Arq' (text), source data.csv\n- Power Query: 1 connection (the queries are in ## Queries)\n"), cons
    assert not any(k in txt for k in ("ana", "hunter", "abc9", "zzz9", "Users")), txt
    assert "- connection 'Vendas' refreshes when the file opens" in txt and "'## Connections' (after the last sheet)" in txt
    assert "# Workbook" not in converter(_injetar(_fixture(), {"xl/connections.xml": cx}), "Fluxo!D5")[0]
    assert len(conexoes(_injetar(_fixture(), {"xl/connections.xml": cx.replace(b"</connections>", b'<connection id="9" name="n"/>' * 100 + b"</connections>")}))[0]) == 102
    grande = b"<connections>" + b" " * (6 * 2 ** 20) + b"</connections>"
    assert "skipped" in conexoes(_injetar(_fixture(), {"xl/connections.xml": grande}))[1][0]
    assert "skipped" in conexoes(_injetar(_fixture(), {"xl/connections.xml": b'<!DOCTYPE a [<!ENTITY e "x">]><a/>'}))[1][0]
    assert "skipped" in conexoes(_injetar(_fixture(), {"xl/connections.xml": b"<a>"}))[1][0]
    assert "type unknown" in conexoes(_injetar(_fixture(), {"xl/connections.xml": b'<connections><connection id="1" name="a" type="\xc2\xb2"/></connections>'}))[0][0]
    assert conexoes(_fixture()) == ([], []) and conexoes(io.BytesIO(b"nao e zip")) == ([], [])
    # power query: blob sintetico (base64 -> cabecalho -> zip interno -> Section1.m + metadados)
    def _mashup(m_texto, meta_xml, codif="utf-8-sig"):
        bz = io.BytesIO()
        with zipfile.ZipFile(bz, "w") as zq:
            zq.writestr("Formulas/Section1.m", m_texto.encode(codif))
        pacote = bz.getvalue()
        xml = meta_xml.encode("utf-8")
        meta = struct.pack("<II", 0, len(xml)) + xml + struct.pack("<I", 0)
        blob = struct.pack("<II", 0, len(pacote)) + pacote + struct.pack("<I", 0) + struct.pack("<I", len(meta)) + meta + struct.pack("<I", 0)
        b64 = base64.b64encode(blob).decode()
        b64 = b64[:20] + "\n  " + b64[20:]  # espaco em branco dentro do base64
        return ('<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">%s</DataMashup>' % b64).encode("utf-16")
    m_txt = ('section Section1;\r\n\r\nshared Vendas = let\r\n    Source = Excel.Workbook(File.Contents("C:\\Users\\ana\\x.xlsx")),\r\n'
             '    Web = Web.Contents("https://h.com/p?sig=zzz9", [Headers=[Authorization="Bearer abc9"]]),\r\n    Conn = [Password = "hunter2"]\r\nin\r\n    Source;\r\n\r\n'
             'shared #"Vendas 2024" = let\r\n    S = Vendas\r\nin\r\n    S;\r\n')
    meta_x = ('<LocalPackageMetadataFile><Items><Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Vendas</ItemPath></ItemLocation>'
              '<StableEntries><Entry Type="FillEnabled" Value="l1"/><Entry Type="FillTarget" Value="sVendas"/><Entry Type="FillToDataModelEnabled" Value="l1"/></StableEntries></Item>'
              '<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Vendas%202024</ItemPath></ItemLocation>'
              '<StableEntries><Entry Type="FillEnabled" Value="l0"/></StableEntries></Item></Items></LocalPackageMetadataFile>')
    cm = _injetar(_fixture(), {"customXml/item1.xml": _mashup(m_txt, meta_x)})
    lista, ok = consultas(cm)
    assert ok and [q[0] for q in lista] == ["Vendas", "Vendas 2024"] and lista[1][1].startswith("let") and lista[1][1].endswith("S"), lista
    assert [q[2] for q in lista] == ["sheet table Vendas, data model", "connection only"]
    txt, _ = converter(cm)
    qs = txt.split("\n# Workbook\n")[1]
    assert qs.startswith("## Queries\n- Vendas, loads to: sheet table Vendas, data model\n    let\n        Source = Excel.Workbook(File.Contents(\"x.xlsx\")),\n"), qs
    assert "- Vendas 2024, loads to: connection only\n    let\n        S = Vendas\n    in\n        S" in qs
    assert "https://h.com/p" in qs and "Password=***" in qs and "Authorization=***" in qs
    assert not any(k in txt for k in ("ana", "zzz9", "abc9", "hunter2", "Users")), txt
    assert "'## Queries' (after the last sheet)" in txt and "# Workbook" not in converter(cm, "Fluxo!D5")[0]
    # truncamento, limite, destino ausente, sem Power Query, blob quebrado
    longo = _injetar(_fixture(), {"customXml/item1.xml": _mashup("section Section1;\nshared Q = " + "x" * 5000 + ";", "<a/>")})
    sq = secao_consultas(longo)
    assert sq[0] == "- Q" and sq[1] == "    " + "x" * LIMITE_M + "...(truncated)", sq
    muitas = "section Section1;\n" + "".join("shared Q%d = 1;\n" % i for i in range(60))
    assert secao_consultas(_injetar(_fixture(), {"customXml/item1.xml": _mashup(muitas, "<a/>")}))[-1] == "(+10 more queries)"
    assert consultas(_injetar(_fixture(), {"customXml/item1.xml": _mashup("section Section1;", "<a/>")})) == ([], True)
    assert consultas(_fixture()) == ([], True) and consultas(io.BytesIO(b"nao zip")) == ([], True)
    assert consultas(_injetar(_fixture(), {"customXml/item1.xml": b""})) == ([], True)
    assert consultas(_injetar(_fixture(), {"customXml/item1.xml": b"<a>" + b" " * (6 * 2 ** 20) + b"</a>"})) == ([], True)
    for ruim in (b'<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">AAAA</DataMashup>', b'<DataMashup>!!nao-base64</DataMashup>',
                 b'<!DOCTYPE DataMashup [<!ENTITY e "x">]><DataMashup>&e;</DataMashup>'):
        assert secao_consultas(_injetar(_fixture(), {"customXml/item1.xml": ruim})) == ["Queries present but unreadable"]
    assert "## Queries" not in converter(_injetar(_fixture(), {"customXml/item1.xml": b"<outra/>"}))[0]
    # macros: container CFB e compressao OVBA sinteticos (o mesmo codigo foi validado a mao num .xlsm real, nao versionado)
    def _cfb(fluxos):
        ss, FIM, LIVRE, FATSEC = 512, 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD
        pastas = sorted({c.rpartition("/")[0] for c in fluxos if "/" in c})
        ent = [["Root Entry", 5, b"", 0]] + [[p, 1, b"", 0] for p in pastas]
        for c, d in sorted(fluxos.items()):
            pai, _, nome = c.rpartition("/")
            ent.append([nome, 2, d, pastas.index(pai) + 1 if pai else 0])
        fat, setores, inicio = [FATSEC], [], {}  # o setor 0 e a FAT

        def aloca(dados):
            n = -(-len(dados) // ss)
            ini = len(setores) + 1
            for k in range(n):
                setores.append(dados[k * ss:(k + 1) * ss].ljust(ss, b"\0"))
                fat.append(ini + k + 1 if k < n - 1 else FIM)
            return ini if n else FIM
        ndir = -(-len(ent) * 128 // ss)
        dir0 = aloca(b"\0" * ndir * ss)
        mini, minifat = b"", []
        for i, e_ in enumerate(ent):
            if e_[1] == 2 and 0 < len(e_[2]) < 4096:  # pequeno: vai para o mini stream (setores de 64 bytes)
                inicio[i] = len(minifat)
                n = -(-len(e_[2]) // 64)
                minifat += [len(minifat) + k + 1 if k < n - 1 else FIM for k in range(n)]
                mini += e_[2].ljust(n * 64, b"\0")
        mini0 = aloca(struct.pack("<%dI" % len(minifat), *minifat)) if minifat else FIM
        raiz_ini = aloca(mini)
        for i, e_ in enumerate(ent):
            if e_[1] == 2 and i not in inicio:
                inicio[i] = aloca(e_[2]) if e_[2] else FIM
        filhos = {}
        for i, e_ in enumerate(ent[1:], 1):
            filhos.setdefault(e_[3], []).append(i)
        direita = {a: b for lst in filhos.values() for a, b in zip(lst, lst[1:])}
        raw = bytearray()
        for i, e_ in enumerate(ent):
            nome = e_[0].encode("utf-16-le")
            ini, tam = (raiz_ini, len(mini)) if i == 0 else (inicio.get(i, 0), len(e_[2]))
            filho = filhos[i][0] if i in filhos else LIVRE
            raw += struct.pack("<64sHBBIII16sI16sIQ", nome, len(nome) + 2, e_[1], 1, LIVRE, direita.get(i, LIVRE), filho, b"", 0, b"", ini, tam)
        for k in range(ndir):
            setores[dir0 - 1 + k] = bytes(raw[k * ss:(k + 1) * ss]).ljust(ss, b"\0")
        cab = struct.pack("<8s16sHHHHH6sIIIIIIIII", CFB_ASSINATURA, b"", 0x3E, 3, 0xFFFE, 9, 6, b"", 0, 1, dir0, 0, 4096,
                          mini0, -(-len(minifat) * 4 // ss) if minifat else 0, FIM, 0) + struct.pack("<109I", 0, *([LIVRE] * 108))
        assert len(fat) <= 128
        return cab + struct.pack("<128I", *(fat + [LIVRE] * (128 - len(fat)))) + b"".join(setores)

    def _ovba_bytes(d):  # compressao so com literais (blocos de 3000 bytes), o bastante para a ida e volta
        saida = bytearray([1])
        for i in range(0, max(len(d), 1), 3000):
            cmp_ = bytearray()
            for j in range(i, min(i + 3000, len(d)), 8):
                cmp_ += b"\0" + d[j:j + 8]
            saida += struct.pack("<H", 0xB000 | (len(cmp_) + 2 - 3)) + cmp_
        return bytes(saida)
    _ovba = lambda texto: _ovba_bytes(texto.encode("cp1252"))

    def _dir(mods):
        r = lambda i, d=b"": struct.pack("<HI", i, len(d)) + d
        b = r(3, struct.pack("<H", 1252)) + r(9, struct.pack("<I", 1)) + b"\0\0" + r(0x0F, struct.pack("<H", len(mods)))
        for nome, off, tipo in mods:
            n = nome.encode("cp1252")
            b += r(0x19, n) + r(0x1A, n) + r(0x31, struct.pack("<I", off)) + r(tipo) + r(0x2B)
        return _ovba_bytes(b + r(0x10))

    def _cifra(dados, semente=0x1F, chave=0x55):  # inverso de _decifra_projeto
        plano = b"\0" * ((semente & 6) >> 1) + struct.pack("<I", len(dados)) + dados
        u1, e1, e2, saida = chave, chave ^ semente, 2 ^ semente, bytearray([semente, 2 ^ semente, chave ^ semente])
        for p in plano:
            c = p ^ ((e2 + u1) & 0xFF)
            saida.append(c)
            e2, e1, u1 = e1, c, p
        return saida.hex().upper()

    def _vba(modulos, protegido=False, extra=None):
        """modulos: [(nome, tipo 0x21|0x22, codigo)] -> bytes de um vbaProject.bin sintetico (7 bytes de 'cache' antes do codigo)."""
        fl = {"VBA/_VBA_PROJECT": b"\1\2", "VBA/__SRP_0": b"zz", "VBA/dir": _dir([(n, 7, t) for n, t, _ in modulos])}
        for n, t, c in modulos:
            fl["VBA/" + n] = b"\xaa" * 7 + _ovba(c)
        cmg, dpb = (_cifra(b"\1\0\0\0"), _cifra(b"\1")) if protegido else (_cifra(b"\0\0\0\0"), _cifra(b"\0"))
        fl["PROJECT"] = ('ID="{X}"\r\nModule=%s\r\nDocument=Plan1/&H00000000\r\nCMG="%s"\r\nDPB="%s"\r\n' % (modulos[0][0], cmg, dpb)).encode("latin-1")
        fl.update(extra or {})
        return _cfb(fl)

    # ovba: exemplo feito a mao (3 literais + 1 copia de 9 bytes a 3 de distancia) e ida e volta
    assert ovba_descomprimir(bytes.fromhex("0105b008616263" "0620")) == b"abcabcabcabc"
    longo = "Sub A()\r\n" + "x = 1 'comentario\r\n" * 500
    assert ovba_descomprimir(_ovba(longo)).decode("cp1252") == longo and len(_ovba(longo)) > 8000
    for ruim in (b"", bytes.fromhex("020005b0"), bytes.fromhex("0105b00861626306"),  # sem assinatura; copia cortada
                 bytes.fromhex("0105b008616263ffff"), bytes.fromhex("0102b0010000")):  # distancia alem do inicio; copia sem nada antes
        try:
            ovba_descomprimir(ruim)
            assert False, ruim
        except ValueError:
            pass
    try:
        ovba_descomprimir(_ovba("a" * 2000), maximo=100)
        assert False, "teto nao aplicado"
    except ValueError:
        pass
    # cfb: ida e volta com fluxo pequeno (mini stream), grande (FAT), aninhado, na raiz e vazio
    flx = {"VBA/dir": b"d" * 100, "VBA/Mod1": bytes(range(256)) * 20, "PROJECT": b"P=1", "VBA/vazio": b""}
    bom = _cfb(flx)
    assert cfb_streams(bom) == flx, cfb_streams(bom).keys()
    for ruim in (b"nao e cfb", bom[:600], bom[:512 * 3], bom[:72] + struct.pack("<I", 1) + bom[76:]):
        try:
            cfb_streams(ruim)
            assert False, "container ruim aceito"
        except ValueError:
            pass
    ciclo = bytearray(bom)
    ciclo[512 + 4:512 + 8] = struct.pack("<I", 1)  # FAT[1] aponta para si mesmo (primeiro setor do diretorio)
    try:
        cfb_streams(bytes(ciclo))
        assert False, "ciclo aceito"
    except ValueError:
        pass
    # entrada hostil: orcamento de descompactacao entre modulos (varios apontando para o mesmo fluxo bomba), regex nao quadratica
    import time
    teto, globals()["MAX_DESCOMPACTADO"] = MAX_DESCOMPACTADO, 10
    try:
        orc = "\n".join(secao_macros(_injetar(_fixture(), {"xl/vbaProject.bin": _vba([("A1", 0x21, "Sub A()\r\nEnd Sub"), ("A2", 0x21, "Sub B()\r\nEnd Sub")])}), True))
    finally:
        globals()["MAX_DESCOMPACTADO"] = teto
    assert "- A1 (module): Sub A" in orc and "- A2 (module): source unreadable" in orc, orc
    t0 = time.time()
    RE_ESCRITA.findall("\n" * 200000 + "a(ba(b" * 50000)
    RE_PROCEDIMENTO.findall("\n" * 200000)
    assert time.time() - t0 < 5, "regex quadratica"
    # protecao: CMG/DPB sempre existem; sem senha decifram para zero (valores iguais aos de um arquivo real sem senha)
    assert not projeto_protegido('CMG="1F1D3D73C5F315F715F715F715F7"\r\nDPB="6B6949C794C894C894"\r\nGC="B7B5959BADABF8ACF8AC07"')
    assert projeto_protegido('CMG="%s"\r\nDPB="%s"' % (_cifra(b"\0\0\0\0"), _cifra(b"\x01\x02")))
    assert projeto_protegido('CMG="%s"\r\nDPB="%s"' % (_cifra(b"\1\0\0\0"), _cifra(b"\0")))
    assert projeto_protegido('CMG="ZZ"\r\nDPB="00"') and not projeto_protegido("sem nada")
    # ## Macros ponta a ponta: nomes por padrao, codigo so com a opcao; segredos plantados somem, o resto fica
    cod1 = ('Attribute VB_Name = "Modulo1"\r\nSub Gravar()\r\n    Sheets("A").Range("B2").Value = 5\r\n    Cells(1, 2) = 7\r\n    Range("B2").Value = 6\r\n'
            '    If Range("C1") = 3 Then x = 1\r\n    .Value2 = 9\r\n    pwd = "hunter2"\r\n    Workbooks.Open "C:\\Users\\ana\\dados.xlsx"\r\nEnd Sub\r\n'
            'Public Function Soma(a)\r\n    Soma = a\r\nEnd Function\r\n')
    mods = [("Modulo1", 0x21, cod1), ("Plan1", 0x22, "x" * 6000), ("M\u00f3dulo2", 0x21, "Sub Z()\r\nEnd Sub\r\n")]
    xm = _injetar(_fixture(), {"xl/vbaProject.bin": _vba(mods)})
    nomes_m = secao_macros(xm)
    assert nomes_m[1] == "(source not shown: run with --macros-fonte)" and len(nomes_m) == 2, nomes_m
    assert set(nomes_m[0][len("modules: "):].split(", ")) == {"Modulo1", "Plan1", "M\u00f3dulo2"}, nomes_m
    full = "\n".join(secao_macros(xm, True))
    assert "- Modulo1 (module): Sub Gravar, Function Soma" in full and "- M\u00f3dulo2 (module): Sub Z" in full, full
    assert 'writes to cells:\n    Sheets("A").Range("B2").Value\n    Cells(1, 2)\n    Range("B2").Value\n    .Value2\n  source:' in full, full
    assert full.count("writes to cells:\n    Sheets") == 1 and "Attribute VB_Name" not in full and 'If Range("C1") = 3' in full
    assert "dados.xlsx" in full and not any(k in full for k in ("hunter2", "ana", "Users")), full
    assert "- Plan1 (document)" in full and "x" * 4000 + "...(truncated)" in full and "x" * 4001 not in full
    muitos = [("Mod%d" % i, 0x21, "y%d = 1 " % i * 800) for i in range(8)]
    tot = "\n".join(secao_macros(_injetar(_fixture(), {"xl/vbaProject.bin": _vba(muitos)}), True))
    assert "source omitted: total cap of 20000" in tot and tot.count("...(truncated)") >= 1 and len(tot) < 24000
    prot = secao_macros(_injetar(_fixture(), {"xl/vbaProject.bin": _vba(mods, protegido=True)}), True)
    assert prot[0].startswith("- VBA project is password protected") and prot[1].startswith("modules: ") and len(prot) == 2, prot
    # modulo com fluxo corrompido: so ele perde o codigo; dir ilegivel: so os nomes
    quebrado = _vba(mods, extra={"VBA/Modulo1": b"\xaa" * 7 + b"\x01\xff\xff"})
    assert "- Modulo1 (module): source unreadable" in "\n".join(secao_macros(_injetar(_fixture(), {"xl/vbaProject.bin": quebrado}), True))
    sem_dir = secao_macros(_injetar(_fixture(), {"xl/vbaProject.bin": _vba(mods, extra={"VBA/dir": b"\x00"})}), True)
    assert sem_dir[-1] == "(project index unreadable: no source)" and sem_dir[0].startswith("modules: ")
    # R5: sem VBA, ou vbaProject.bin que nao e CFB: sem ## Macros, so o aviso antigo
    assert secao_macros(_fixture()) == [] and secao_macros(io.BytesIO(b"nao zip")) == []
    falso = converter(_injetar(_fixture(), {"xl/vbaProject.bin": b"x"}))[0]
    assert "## Macros" not in falso and "File has VBA macros" in falso
    txt, _ = converter(xm)
    assert txt.split("\n# Workbook\n")[1].startswith("## Macros\nmodules: ") and "'## Macros' (after the last sheet)" in txt and "Sub Gravar" not in txt
    assert "Sub Gravar" in converter(xm, None, True)[0] and "# Workbook" not in converter(xm, "Fluxo!D5")[0]
    with tempfile.TemporaryDirectory() as td:  # a opcao pela linha de comando
        caminho = os.path.join(td, "m.xlsm")
        with open(caminho, "wb") as f_:
            f_.write(xm.getvalue())
        sys.stdout, saida_ = io.StringIO(), sys.stdout
        try:
            assert main(caminho, None, True) == 0
        finally:
            sys.stdout = saida_
        assert "Sub Gravar" in open(os.path.join(td, "m.ai.txt"), encoding="utf-8").read()
    print("auto-teste ok")


if __name__ == "__main__":
    macros_fonte = "--macros-fonte" in sys.argv
    if macros_fonte:
        sys.argv.remove("--macros-fonte")
    if len(sys.argv) == 4 and sys.argv[2] == "--celula":
        sys.exit(main(sys.argv[1], sys.argv[3], macros_fonte))
    if len(sys.argv) == 4 and sys.argv[3] == "--diff":
        sys.exit(main_diff(sys.argv[1], sys.argv[2]))
    if sys.argv[1:2] == ["--instalar-skill"] and len(sys.argv) <= 3:
        sys.exit(instalar_skill(forcar=sys.argv[2:] == ["--forcar"]))
    if len(sys.argv) == 2:
        sys.exit(main_pasta(sys.argv[1], macros_fonte) if os.path.isdir(sys.argv[1]) else main(sys.argv[1], None, macros_fonte))
    if len(sys.argv) > 1:
        sys.stderr.write("uso: python excel_ai.py planilha.xlsx [--celula Aba!B5] [--macros-fonte]  |  python excel_ai.py antiga.xlsx nova.xlsx --diff  |  python excel_ai.py pasta  |  python excel_ai.py --instalar-skill [--forcar]\n")
        sys.exit(2)
    autoteste()
