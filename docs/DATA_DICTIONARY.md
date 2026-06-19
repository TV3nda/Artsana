# Dicionario De Dados

Fonte principal: `data/historico/wells_historico.csv`.

Cada linha representa a observacao de um produto num dia.

| Campo | Significado |
| --- | --- |
| `Data` | Dia da recolha, formato `YYYY-MM-DD`. |
| `Hora` | Hora da recolha. |
| `Categoria` | Categoria monitorizada: `Chupetas`, `Biberoes`, `Bombas_Tira_Leite`. |
| `ProdID` | Identificador do produto na Wells. Chave principal junto com `Data`. |
| `Marca` | Marca do produto. |
| `Produto` | Nome do produto no site Wells. |
| `Preco` | Preco atual observado. |
| `PVPR` | Preco original quando ha desconto. Vazio quando nao ha diferenca face ao preco atual. |
| `Desconto_Pct` | Percentagem de desconto calculada. |
| `Poupanca_Euro` | Diferenca entre PVPR e preco atual. |
| `Destaque` | Badges observados no tile, como `Best Seller`, `Exclusivo Online` ou `Novo`. |
| `Stock` | `Disponivel` ou `Sem Stock`, segundo sinais do site. |
| `URL` | Link do produto. |

## Metricas Derivadas No Dashboard

| Metrica | Definicao |
| --- | --- |
| Novidade | Produto cujo primeiro dia observado no historico cai dentro do mes selecionado. |
| Referencia em promocao | Produto com `Desconto_Pct` preenchido pelo menos uma vez no periodo. |
| Dias em promocao | Numero de dias observados com desconto. |
| Primeira promocao observada | Produto com promocao no mes e sem promocao anterior no historico disponivel. |
| Ausente no fim do periodo | Produto observado durante o mes mas sem registo no ultimo dia observado desse mes. E apenas sinal online. |
| Intensidade promocional | Percentagem de referencias observadas que estiveram em promocao. |

## Cuidados De Interpretacao

- Dados do site Wells nao equivalem automaticamente ao cardex de lojas fisicas.
- Ausencia online pode significar remocao, indisponibilidade, alteracao de URL/nome ou falha de recolha.
- Impacto de mercado deve ser confirmado com sell-out.
- Nomes de produtos podem mudar; quando isso acontecer, `ProdID` deve prevalecer.

