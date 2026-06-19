# Roadmap

## Feito nesta fase

- Consolidacao da versao GitHub como fonte principal.
- Dashboard mensal com novidades, promocoes, presenca online e texto base para email.
- Comparacao Chicco vs mercado Wells online.
- Exportacao mensal para Excel.
- Protecao contra duplicados futuros no historico.
- Deduplicacao do historico existente por `Data + ProdID`.
- Validacao basica para evitar gravar snapshots incompletos.
- Documentacao para humanos e outras AIs.

## Proxima fase recomendada

1. **Alertas semanais**
   - resumo de novidades;
   - promocoes fortes;
   - produtos que desapareceram no fim da semana.

2. **Camada de curadoria humana**
   - marcar novidade como relevante/irrelevante;
   - adicionar notas comerciais;
   - guardar comentarios para o relatorio mensal.

3. **Integracao com sell-out**
   - importar ficheiro mensal de sell-out;
   - cruzar sinais Wells com performance real por categoria/marca;
   - separar "sinal observado" de "impacto confirmado".

4. **Qualidade e testes**
   - testes automaticos do parser;
   - alerta quando a Wells muda HTML;
   - validacao visual do dashboard.

5. **Melhorias de dashboard**
   - filtros mensais mais finos;
   - graficos de intensidade promocional;
   - historico de preco minimo por produto;
   - ranking de marcas por novidades e promocoes.

