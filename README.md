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

Para olhar uma conta só (por exemplo a TIR final), passe a célula:

```
python excel_ai.py sua_planilha.xlsx --celula "Resumo!B5"
```

Gera `sua_planilha.Resumo_B5.ai.txt` só com as fórmulas e entradas de que essa célula depende, em todas as abas.

Sem argumento, roda o auto-teste:

```
python excel_ai.py
```
