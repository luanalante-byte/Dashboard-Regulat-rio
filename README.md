# Dashboard Regulatório — Pronutrition

Rascunho de dashboard de indicadores da equipe Regulatório, gerado a partir da planilha
**LISTA MESTRE DE PRODUTOS - PRONUTRITION.xlsx** (SharePoint/OneDrive). Publicado
automaticamente em: https://dashboard-regulatory.netlify.app/

## Estrutura

```
dashboard/
  index.html                     — dashboard interativo (Chart.js), publicado via Netlify
  KPI_Regulatorios_Rascunho.pdf  — versão em PDF dos mesmos indicadores
data/
  kpis_all.json                  — últimos indicadores calculados
```

## Indicadores incluídos

- Funil de cotações (taxa de aprovação, lead time)
- Documentações por tipo (Dizeres de Rotulagem, Correção de Arte, Homologação de MP,
  Avaliação de Risco, Especificação Técnica) com SLA médio e filtro mensal
- Revisões de arte por produto/cliente (distribuição e top 15 com mais retrabalho)
- Gastos com produtos de catálogo — fluxo de caixa mensal por laboratório
- Gastos com clientes — estudos de estabilidade (aba "Gastos Reais")
- Notificações vigentes (por apresentação, marca, cliente, peso líquido e por produto)
- Funil de estabilidade por tempo de estudo e próximos laudos previstos
- Qualidade por laboratório (aprovação / reanálise / reprovação)

## Publicação

O site `dashboard-regulatory.netlify.app` está conectado a este repositório via
deploy contínuo do Netlify (Site settings → Build & deploy → pasta `dashboard`).
Qualquer push na branch `main` publica automaticamente a nova versão.

## Status

Rascunho para validação de layout e indicadores com a equipe Regulatório.
