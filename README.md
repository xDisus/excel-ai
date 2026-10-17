# excel-ai

Transforma uma planilha Excel (calculadoras financeiras, PROCX entre abas, fluxos, TIR) em um texto compacto que uma IA lê rápido e entende.

- Fórmulas repetidas viram uma linha só (`E2:E361 =C2*G2`), com as exceções listadas.
- Cada bloco ganha o nome do cabeçalho mais perto (`E2:E361 [Juros] =C2*G2`), para a IA saber o que a conta é.
- Cada bloco de 2+ celulas mostra o valor que o Excel calculou na primeira celula (`E2:E361 [Juros] (v 12.5) =C2*G2`): so uma dica, a formula nao muda.
- Rótulos de texto ficam inteiros. Tabelas grandes de números viram resumo (faixa, mín, máx, primeiros valores).
- Mapa das abas: quem usa quem, abas ocultas, tabelas, nomes definidos.
- Avisos: `#REF!`, erros, `DESLOC`/`INDIRETO`, macros.
- Conferência automática: o texto é expandido de volta e comparado com cada fórmula original. Se algo não bater, aquela aba sai célula por célula.

Roda 100% local. Nenhum dado sai da sua máquina.

## Instalar

Só precisa de Python 3.8+ e do openpyxl:

```
pip install openpyxl
```

Copie o arquivo `excel_ai.py`. É um arquivo só, sem outras dependências. Link direto para copiar o código:
https://raw.githubusercontent.com/xDisus/excel-ai/main/excel_ai.py

## Usar

```
python excel_ai.py sua_planilha.xlsx
```

Gera `sua_planilha.ai.txt` ao lado da planilha. É esse arquivo que você manda para a IA.
Na tela aparecem só números: tokens antes e depois, e o resultado da conferência.

Se a planilha tem tabelas dinâmicas, o arquivo ganha uma seção `## Pivots` por aba: onde a tabela está, o que a alimenta e quais campos são linhas, colunas, filtros e valores.

```
## Pivots
- 'Piv1' at A3:B8, source 'Dados'!A1:C6
  rows: Mes
  values: Soma de Valor = sum(Valor)
```

Se há links para outras planilhas (`=[1]Taxas!B2`), o final do arquivo ganha `# Workbook` / `## External links`: `[1] taxas.xlsx, sheets: T, 3 formulas`. Só o nome do arquivo aparece, nunca a pasta; o arquivo ligado não é aberto.

Se há conexões de dados (banco, web, arquivo de texto), `## Connections` lista cada uma: tipo, servidor/banco (ou URL/arquivo), comando e se atualiza ao abrir. Usuário, senha e chaves nunca aparecem: da string de conexão só ficam servidor, banco, provedor, driver e fonte de dados.

Para olhar uma conta só (por exemplo a TIR final), passe a célula:

```
python excel_ai.py sua_planilha.xlsx --celula "Resumo!B5"
```

Gera `sua_planilha.Resumo_B5.ai.txt` só com as fórmulas e entradas de que essa célula depende, em todas as abas.

Para ver o que mudou entre duas versões (por exemplo v7.3 e v7.4), passe as duas e `--diff`:

```
python excel_ai.py v7.3.xlsx v7.4.xlsx --diff
```

Gera `v7.4.vs_v7.3.diff.ai.txt` só com o que mudou: abas novas ou removidas, fórmulas novas ou alteradas (com a fórmula antiga ao lado) e constantes alteradas. O que não mudou fica de fora. Se inserir linhas ou colunas, muita coisa aparece como mudança.

Para converter todas as planilhas de uma pasta (sem entrar em subpastas):

```
python excel_ai.py minha_pasta
```

Cada `.xlsx`/`.xlsm` gera o seu `.ai.txt` ao lado. Se uma falhar, as outras seguem.

## Skill do Kiro

Para o Kiro usar a ferramenta sozinho quando você falar de planilha:

```
python excel_ai.py --instalar-skill
```

Grava a skill em `~/.kiro/skills/excel-ai`. Se já existir, pula; `--forcar` sobrescreve. A skill vem de dentro do próprio `excel_ai.py`.

Sem argumento, roda o auto-teste:

```
python excel_ai.py
```
