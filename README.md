# Banco de Horas - Relatório Automatizado por Coligada

## 📌 Descrição

Este projeto automatiza a geração do relatório de **saldo de banco de horas** por colaborador. Elimina a necessidade de copiar e consolidar saldos manualmente no Excel, utilizando **VBA**, **SQL** e integração com a **API do TOTVS RM**.

A automação realiza:

- Extração dos dados via API TOTVS.
- Criação de aba “SALDO COL{n}” para cada coligada.
- Formatação e exibição dos dados salariais e de horas.
- Salvamento automático do arquivo em diretório específico.

## ⚙️ Tecnologias Utilizadas

- **VBA (Visual Basic for Applications)**: lógica de extração e criação das planilhas.
- **SQL (TOTVS / RM Reports)**: consultas para saldo.
- **Excel**: exibição e estrutura do relatório.
- **API TOTVS RM**: origem dos dados.

## 🧠 Lógica da Automação

1. Um painel em Excel executa um loop para cada coligada configurada.
2. A Sub `Extrair_API_Nova` executa:
   - Requisição HTTP à API.
   - Processamento do JSON.
   - Criação da aba `SALDO COL{n}` com os dados estruturados.
   - Inclusão de bordas, cores e cabeçalhos padronizados.
3. O arquivo é salvo automaticamente com nome personalizado, por coligada e período.

## 📁 Estrutura dos Arquivos

BancoHoras/
├── VBA/
│ └── Extrair_API_Nova.bas
├── SQL/
│ └── ConsultaBancoHoras.sql
├── README.md
└── ExemploRelatorio/


## ✅ Resultados

- Redução de trabalho mensal de **3 a 5 horas** para **menos de 5 minutos**.
- Dados centralizados por coligada, com padronização visual automática.

## 🚧 Melhorias Futuras

- Incluir envio automático por e-mail.
- Criar interface com botões no Excel.
- Aplicar criptografia ao salvar relatórios com dados sensíveis.

- ## 👤 Autor

**Marcus Vinícius da Silva Nunes**  
Analista de Departamento Pessoal em transição para a área de Tecnologia.

- 💼 [LinkedIn](https://www.linkedin.com/in/marcus-vinicius-da-silva-nunes-01b784125/)
