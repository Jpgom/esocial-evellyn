
E-SOCIAL EVELLYN - VERSÃO FINAL V1

REGRAS DESTA VERSÃO
- a planilha base agora usa:
  FUNCIONÁRIO
  TIPO DE EXAME
  DEPOSITANTE
- a planilha do sistema continua usando:
  NOME
  TIPO

REGRA NOVA PARA GERAR PDF
- o PDF de uma empresa só é gerado se TODOS os funcionários daquela empresa na planilha base estiverem com OK E-SOCIAL na coluna DEPOSITANTE
- além disso, os funcionários esperados também precisam ser encontrados no RELFUNCGERAL

SAÍDA FINAL
- PDFs/
- Logs/
- RESUMO_FINAL.xlsx

O arquivo RESUMO_FINAL.xlsx informa:
- empresa
- cnpj
- se gerou ou não gerou PDF
- quantidade esperada na base
- quantidade encontrada no sistema
- motivo
- caminho do PDF gerado
