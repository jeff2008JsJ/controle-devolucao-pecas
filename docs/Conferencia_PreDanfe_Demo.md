# Conferência de Pré-DANFE com VBA — versão de portfólio

Este projeto demonstra uma automação em Excel/VBA para apoiar a conferência de documentos fiscais em PDF recebidos por e-mail, usando o Outlook e tabelas de referência do Excel.

## O que o módulo demonstra

- Seleção de e-mail no Outlook e identificação de PDFs/XMLs anexados.
- Extração de texto do PDF e alternativa de leitura por meio do Word quando necessário.
- Conferência de dados da NF-e contra cadastros e pendências da planilha.
- Validação de campos fiscais e identificação de divergências para revisão.
- Registro dos resultados em abas de conferência e histórico.
- Separação entre conferência do documento e baixa definitiva dos chamados.

## Privacidade e dados de demonstração

Esta versão de portfólio foi anonimizada. O nome da empresa, CNPJ, e-mails, nomes de transportadoras, links e identificadores de STA foram substituídos por valores fictícios ou reservados para exemplos. O CNPJ demonstrativo é inválido e não deve ser usado em operações reais.

## Avisos importantes

- O módulo depende da estrutura das planilhas e das colunas esperadas pelo código. Ele não funciona isoladamente sem uma pasta de trabalho compatível.
- As regras fiscais e operacionais incluídas são exemplos de implementação e precisam ser revisadas por profissionais responsáveis antes de qualquer uso real.
- O projeto não substitui a validação fiscal oficial nem a autorização da NF-e pela SEFAZ.
- Antes de testar macros que interagem com o Outlook, use dados fictícios e revise os destinatários.
- Não coloque no repositório documentos fiscais, e-mails, CNPJs, nomes de parceiros ou dados internos reais.

## Tecnologias

- Microsoft Excel e VBA
- Microsoft Outlook Object Model
- Leitura de PDF/XML e tratamento de texto
- Dicionários e validação cruzada de dados

