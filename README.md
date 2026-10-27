# AgTech Demand Radar

Radar automatizado de sinais públicos de investimento e inovação na agroindústria brasileira.

## Executar

Na aba **Actions**, abra **AgTech Demand Radar** e clique **Run workflow**. Execução semanal programada: segunda-feira, 07:20 UTC.

O workflow consulta BNDES, Finep e Embrapii, salva um artefato de execução e prepara um painel estático (`site/`). Ele **não** interpreta financiamentos como prova de disposição de compra.

## Dashboard

Acesse **Settings → Pages → Build and deployment → Source: GitHub Actions** para habilitar o Pages (conforme disponibilidade do seu plano para repositórios privados). O site é estático e não contém credenciais. Executar o pipeline diretamente por um botão no Pages exigiria autenticação/backend adicional; execute pelo GitHub Actions.

## Limitações

As integrações de dados são exploratórias e dependem de estruturas externas que podem mudar. `status.json` registra falhas; frequência de temas não é ranking validado de oportunidades de mercado. A primeira execução online precisa ser verificada.
