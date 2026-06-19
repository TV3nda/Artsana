# Wells Price Intelligence

Dashboard e scraper para acompanhar precos, novidades, promocoes e sinais de presenca online na Wells.pt para categorias de bebe.

## Estado atual

Esta pasta e a versao canonica do projeto:

```text
C:\AI Workspace\Dashboard Wells\github\Artsana
```

As copias antigas devem ficar arquivadas fora desta pasta. Nao usar `wells_scraper`, `$repo` ou `backups` como fonte de desenvolvimento.

## Para que serve

O projeto recolhe dados diarios da Wells.pt para:

- Chupetas
- Biberoes
- Bombas tira-leite

Depois gera um dashboard estatico com:

- vista diaria de produtos;
- comparacao entre datas;
- evolucao de precos por produto;
- vista por marca;
- relatorio mensal com novidades, intensidade promocional, sinais de presenca online e insights para apoiar o email mensal.

O dashboard publicado vive em:

```text
https://tv3nda.github.io/Artsana
```

## Como funciona

1. `scripts/wells_scraper.ps1` chama o endpoint Ajax da Wells.
2. Extrai produto, marca, preco, PVPR, desconto, stock, destaque e URL.
3. Guarda um CSV diario em `data/recente/`.
4. Atualiza `data/historico/wells_historico.csv` como snapshot do dia, evitando duplicados.
5. `scripts/gerar_dashboard.ps1` le o historico e gera `index.html`.
6. O GitHub Actions corre isto diariamente e faz commit do historico e do dashboard.

## Comandos uteis

Gerar o dashboard localmente a partir do historico:

```powershell
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass `
  -File scripts/gerar_dashboard.ps1 `
  -DataDir data `
  -DashboardOut index.html
```

Correr o scraper localmente:

```powershell
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass `
  -File scripts/wells_scraper.ps1 `
  -OutputDir data
```

## Ficheiros principais

- `scripts/wells_scraper.ps1`: recolha diaria e validacao.
- `scripts/gerar_dashboard.ps1`: fonte do dashboard HTML.
- `index.html`: dashboard gerado para GitHub Pages.
- `data/historico/wells_historico.csv`: historico acumulado.
- `.github/workflows/scraper.yml`: automacao diaria no GitHub Actions.

## Regras importantes

- Nao editar `index.html` manualmente para alterar funcionalidades. Editar `scripts/gerar_dashboard.ps1` e regenerar.
- Antes de mexer no historico, criar backup.
- O campo de presenca/ausencia e leitura do site Wells, nao confirmacao de cardex em loja fisica.
- O texto mensal e uma base editavel; deve ser cruzado com sell-out e conhecimento de mercado antes de enviar.
- Nao guardar API keys ou segredos no repositorio.
- Nao colocar sell-out, relatorios internos ou documentos confidenciais neste repositorio publico. Se for preciso usar esses ficheiros para uma analise local, guardar em `private/`, `private_data/`, `sellout_private/` ou `local_private/`, que estao ignoradas pelo Git.

## Dados Publicos Vs Dados Privados

Este dashboard publico deve conter apenas dados recolhidos do site Wells.pt e outros dados que possam ser publicados sem risco.

Dados privados, como sell-out, relatorios internos, apresentacoes comerciais ou notas confidenciais, devem ficar fora do GitHub e fora do `index.html`. A forma segura de os usar e fazer uma analise local separada, gerar conclusoes privadas, e depois decidir manualmente que frases podem entrar no email mensal.

## Validacao recomendada

Depois de alterar scripts:

1. Gerar `index.html` localmente.
2. Confirmar que nao ha duplicados por `Data + ProdID`.
3. Abrir o dashboard e testar as abas principais.
4. Confirmar que a aba `Relatorio Mensal` gera texto e tabelas.
