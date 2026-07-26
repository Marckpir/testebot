# testebotaprweb-node

Clone em Node.js da aplicação APR com Playwright.

## Funcionalidades

- Consulta NS individual (dados gerais + anexos + históricos)
- Consulta em lote (retorna `NS` + `NOME RESPONSAVEL TECNICO`)
- Limpeza do nome no lote (somente letras e espaços)
- Sem leitura de anexos/históricos no modo lote
- Campo de tipo de pesquisa com exibição condicional de campos
- Exporta `NUMERO DE INSTALAÇÃO` da consulta individual para `localStorage` na chave `aprweb.numero_instalacao`

## Requisitos

- Node.js 18+

## Instalação

```bash
npm install
npx playwright install chromium
```

## Execução

```bash
npm start
```

Acesse: http://127.0.0.1:5050

## Implantação no Vercel

1. Crie um projeto no Vercel e conecte este repositório.
2. O projeto já inclui um endpoint serverless em `api/index.js` e um arquivo `vercel.json` para encaminhar todas as rotas para ele.
3. Defina as variáveis de ambiente se necessário:
   - `APR_USUARIO`
   - `APR_SENHA`
4. Deploy no Vercel.

> Importante: a automação de navegador depende de um ambiente com suporte a Chrome. No Vercel, o código tenta usar `chrome-aws-lambda` para localizar o binário compatível com serverless.
