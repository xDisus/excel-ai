# Skill Kiro: excel-ai

Ensina o Kiro a ler planilhas Excel pelo `excel_ai.py` (texto compacto, sem perda) em vez de abrir o arquivo cru.

## Instalar

Global (todos os projetos):

```
cp -r kiro/skills/excel-ai ~/.kiro/skills/
```

Só neste projeto: copie a pasta para `.kiro/skills/excel-ai` na raiz do projeto.

Uma vez: `pip install openpyxl`.

O Kiro carrega a skill sozinho quando você fala de planilha, fórmula, PROCX, TIR etc. Agente personalizado: adicione `skill://excel-ai` em `resources`.

`scripts/excel_ai.py` é uma cópia do `excel_ai.py` da raiz do repo. Para atualizar: `cp excel_ai.py kiro/skills/excel-ai/scripts/excel_ai.py`. O auto-teste avisa se as duas divergirem.
