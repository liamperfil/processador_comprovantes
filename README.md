# 🧾 Projeto: Processador de Comprovantes

## 🎯 Objetivo
Criar um script Python que:
- Lê arquivos PDF com comprovantes de pagamento;
- Extrai dados (cliente, data de pagamento, valor, código de barras);
- Atualiza uma planilha Excel (`pagamentos.xlsx`);
- Renomeia e move os PDFs para pastas específicas;
- Gera logs com os eventos processados.

## 📂 Estrutura
- **Planilha:** `pagamentos.xlsx`
  - **Abas:** `amoedo`, `beatriz`, `cavalcante`, `dsouza`, `julia`
  - **Colunas:** `id`, `pagamento`, `vencimento`, `beneficiario`, `valor do documento`, `valor cobrado`, `documento`, `codigo de barras`, `categoria`, `status`, `origem`

- **Entrada:** PDFs na raiz do projeto  
- **Saída:** PDFs movidos para subpasta comprovantes/"yyyy-mm"/  
- **Logs:** gerados na pasta `logs` com data  

## ⚙️ Funcionalidades
- ✔ Identificação do cliente por palavra-chave no PDF, priorizando "cliente" ou "empresa"
- ✔ Extração do código de barras em diferentes formatos (com espaços, hifens, etc.) e converter para string com 47 digitos
- ✔ Extração da data de pagamento
- ✔ Extração do valor cobrado
- ✔ Se data de pagamento não for encontrada, usar data de vencimento da planilha
- ✔ Preenchimento do campo "origem": `Bradesco` para `julia`, `Banco do Brasil` para os demais
- ✔ Renomear PDF com base no `id` e mover para a pasta comprovantes/"yyyy-mm"/
- ✔ Solicitação manual de `id` se o código de barras não for localizado
- ✔ Log detalhado e impressão no console
- ✔ Backup da planilha antes de alterações

## Info

## Problemas

## 📁 Arquivo do Script
- `processador_comprovantes.py`

## 🛠 Melhorias implementadas
1. Correções:
 - ✅ Correção de uso incorreto do campo "valor" (prioriza "valor total" ou "valor cobrado")
 - ✅ Correção do erro `strptime()` ao tentar converter `datetime` já formatado
 - ✅ Busca por múltiplas linhas com mesmo código de barras e pagamento vazio
 - ✅ Log explícito se o código de barras não for encontrado no PDF
 - ✅ Data formatada corretamente como `dd/mm/yyyy` sem aspas no Excel
2. Alterar na forma de renomear o arquivo pdf, os PDFs devem ser renomeados para cliente_data_id, exemplo amoedo_20250423_222
3. Alterar local para onde os PDFs processados estão sendo movidos:
 - ✅ Não moveremos mais para os subdiretórios amoedo, beatriz, cavalcante, dsouza e julia.
 - ✅ Usaremos o subdiretório comprovantes, dentro dele outro subdiretório de acordo com o ano e mês, exemplo: comprovantes\2025-04\
4. Correção no input do ID, o mesmo não estava funcionando corretamente.
5. Melhoria na extração do valor cobrado.
6. Formatação no texto quando código de barras não encontrado
7. Correções nas buscas e registro de log:
 - Correção na busca por data de pagamento
 - Correção na buscar por valor cobrado
 - Alerta para documento pdf com mais de uma página
