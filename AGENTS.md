# AI Working Notes

Este projeto e mantido por um utilizador de marketing. Explica trade-offs em linguagem simples e evita assumir que o utilizador conhece desenvolvimento web.

## Fonte canonica

Trabalhar apenas nesta pasta:

```text
C:\AI Workspace\Dashboard Wells\github\Artsana
```

Copias antigas devem estar arquivadas. Nao usar outras pastas como fonte da verdade.

## Regra de edicao

- Para UI/dashboard, editar `scripts/gerar_dashboard.ps1`.
- Para recolha de dados, editar `scripts/wells_scraper.ps1`.
- Regenerar `index.html` depois de alterar `scripts/gerar_dashboard.ps1`.
- Nao editar `index.html` diretamente, exceto para diagnostico muito pontual.
- Nao publicar nem fazer push sem pedido explicito.

## Dados

O CSV principal e:

```text
data/historico/wells_historico.csv
```

Grain esperado: uma linha por `Data + ProdID`.

Antes de reescrever o historico:

1. criar backup em `data/historico/`;
2. confirmar contagens antes/depois;
3. manter backups locais fora do Git.

## Confidencialidade

Este repositorio publica `index.html` no GitHub Pages. Nao inserir dados confidenciais no dashboard, no historico publico ou em qualquer ficheiro versionado.

Dados de sell-out, relatorios internos, ficheiros Excel, PowerPoint, Word ou PDF devem ficar em pastas locais ignoradas pelo Git, como:

```text
private/
private_data/
sellout_private/
local_private/
```

Se o utilizador pedir analise com dados internos, tratar como uma analise privada local e nunca misturar esses dados no dashboard publico sem aprovacao explicita.

## Validações

Comando principal:

```powershell
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts/gerar_dashboard.ps1 -DataDir data -DashboardOut index.html
```

Verificar duplicados:

```powershell
Import-Csv data/historico/wells_historico.csv -Delimiter ';' |
  Group-Object Data,ProdID |
  Where-Object Count -gt 1
```

## Contexto de negocio

O objetivo nao e apenas listar precos. A ferramenta apoia relatórios mensais de marketing, sobretudo:

- novidades identificadas na Wells;
- intensidade promocional por marca;
- comparacao Chicco vs mercado Wells online;
- sinais de presenca online a validar;
- insumos para complementar dados de sell-out.

Nao concluir impacto comercial apenas com dados da Wells online. Usar linguagem prudente: "sinal", "observado online", "a validar", "pode indicar".

## Prioridades futuras

Ver `docs/ROADMAP.md`.
