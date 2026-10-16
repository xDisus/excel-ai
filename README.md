# excel-ai

Deixar planilhas Excel (calculadoras financeiras, fórmulas, PROCX entre abas) fáceis e baratas para IA ler.

## Passo 1: medir se vale a pena

`medir_tokens.py` compara quantos tokens a planilha gasta lida célula por célula vs. com fórmulas repetidas agrupadas.
Roda 100% local e imprime só números, nenhum conteúdo da planilha.

```
pip install openpyxl
python medir_tokens.py sua_planilha.xlsx
```

Sem argumento, roda o auto-teste: `python medir_tokens.py`

Regra: redução de 5x ou mais, vale construir o resto.
