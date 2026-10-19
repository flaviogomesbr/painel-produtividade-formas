# Painel de produtividade — montagem e desmontagem de fôrmas

Site estático que lê `dados/dados.json`. Uma rotina agendada baixa a planilha da
nuvem, regenera esse JSON e publica sozinha.

```
index.html                            o painel
dados/dados.json                      dados que ele consome
scripts/baixar.py                     baixa a planilha (Google Sheets ou Microsoft 365)
scripts/gerar_dados.py                converte a planilha em dados.json
.github/workflows/atualizar-dados.yml agendamento da atualização
```

Custo total: zero. O GitHub Pages hospeda de graça e o GitHub Actions dá minutos
ilimitados em repositório público. Cada atualização consome cerca de um minuto.

---

## Parte 1 — Colocar o painel no ar (10 minutos)

1. Crie uma conta em **github.com**, se ainda não tiver. É grátis.
2. Clique em **New repository**. Nome: `painel-formas`. Marque **Public**.
   Clique em **Create repository**.
3. Na tela seguinte, clique em **uploading an existing file** e arraste **todo o
   conteúdo desta pasta**, mantendo as subpastas. Confira que o arquivo
   `.github/workflows/atualizar-dados.yml` subiu — o navegador às vezes ignora
   pastas que começam com ponto. Se ele não aparecer, use **Add file → Create
   new file** e digite o caminho completo no campo do nome.
4. Clique em **Commit changes**.
5. Vá em **Settings → Pages**. Em *Source*, escolha **Deploy from a branch**.
   Em *Branch*, escolha `main` e a pasta `/ (root)`. Clique em **Save**.
6. Espere dois ou três minutos. O painel estará em:
   `https://SEU-USUARIO.github.io/painel-formas/`

Pronto — qualquer pessoa com o link abre, sem login.

> **Antes de seguir, decida uma coisa.** Repositório público significa que o
> `dados.json` fica acessível a qualquer pessoa na internet: produtividade,
> efetivo e custo por obra. Se isso for sensível, veja a Parte 5.

---

## Parte 2 — Conectar a planilha

O painel aceita as duas nuvens. Escolha o caminho da sua.

### Caminho A — Google Sheets (mais simples)

1. Abra a planilha no Google Sheets.
2. **Compartilhar → Acesso geral → Qualquer pessoa com o link → Leitor.**
3. **Copiar link.**
4. No GitHub: **Settings → Secrets and variables → Actions → New repository
   secret**. Nome: `LINK_PRODUTIVIDADE`. Valor: o link copiado. **Add secret**.
5. Repita para a planilha de custos, com o nome `LINK_CUSTOS`. Esta é opcional —
   sem ela o painel funciona, só a aba de Custo fica sem os valores orçados.

Não precisa de mais nada. A rotina exporta a planilha em xlsx sozinha.

### Caminho B — Excel no OneDrive ou SharePoint

1. Abra a planilha, clique em **Compartilhar → Copiar link**.
2. Guarde nos secrets `LINK_PRODUTIVIDADE` e `LINK_CUSTOS`, como acima.
3. Se o link for do tipo **"Qualquer pessoa com o link"**, acabou.

   Se a empresa bloqueia link anônimo — o caso da maioria dos tenants
   corporativos —, o link só abre para quem está logado, e a rotina precisa de
   credencial própria. Peça ao TI:

   - Portal Azure → **Microsoft Entra ID → App registrations → New registration**.
     Nome: `Painel Formas`. Registrar.
   - **API permissions → Add a permission → Microsoft Graph → Application
     permissions → Files.Read.All** → Add → **Grant admin consent**.
   - **Certificates & secrets → New client secret** → copiar o *Value* na hora
     (ele some depois que você sai da tela).
   - No GitHub, criar mais três secrets: `TENANT_ID` (Directory ID),
     `CLIENT_ID` (Application ID) e `CLIENT_SECRET` (o valor acima).

O script decide sozinho: link do Google usa exportação direta; link da Microsoft
usa o Graph se as credenciais existirem, senão tenta o link anônimo.

---

## Parte 3 — Testar

No GitHub: **Actions → Atualizar dados do painel → Run workflow → Run workflow**.

Em cerca de um minuto o log mostra quantas semanas foram lidas. Recarregue o
painel: o cabeçalho deve exibir **atualização automática** em verde, com a data
da planilha e o horário da geração.

Se aparecer *cópia local* em vermelho, algo falhou. Abra o log da execução em
Actions — a mensagem diz se o problema foi permissão do link ou nome de aba.

---

## Parte 4 — Ritmo da atualização

Já configurado para rodar **de hora em hora, dias úteis, das 7h às 20h de
Brasília**. Para mudar, edite a linha `cron` no arquivo do workflow (o horário é
UTC; Brasília é UTC−3):

| Frequência | cron |
|---|---|
| Uma vez por dia, 7h | `0 10 * * *` |
| A cada 30 min, dias úteis | `*/30 10-23 * * 1-5` |
| A cada 6 horas | `0 */6 * * *` |

O agendamento do GitHub costuma atrasar alguns minutos em horário de pico, e o
commit só acontece quando os dados realmente mudam.

Para atualizar na hora, sem esperar, use o botão **Run workflow**.

---

## Parte 5 — Quando quiser restringir o acesso

GitHub Pages em repositório público é sempre aberto. Para fechar sem reescrever
nada:

1. Crie conta em **cloudflare.com** (plano gratuito).
2. **Workers & Pages → Create → Pages → Connect to Git**, apontando para o mesmo
   repositório.
3. Em **Zero Trust → Access → Applications**, adicione o site e escolha
   **Azure AD / Entra ID** como provedor de login.

Até 50 usuários sem custo. O painel não muda; a autenticação fica por conta do
Cloudflare.

---

## Se você mexer na planilha

O `gerar_dados.py` lê:

- **uma aba por obra**, com nome começando em `Prod_` — `Prod_Serena_CONCLUÍDO`,
  `Prod_Arte-Penha`, `Prod_Origem-Itaquera`, `Prod_Altus_CONCLUÍDO`. Obra nova é
  só criar mais uma aba com esse prefixo; ela aparece no painel sozinha, com cor
  atribuída automaticamente;
- a aba **`GERAL`**, de onde saem o quadro consolidado por obra e (se
  `DATA_ATUALIZACAO_FIXA` estiver vazia) a data de atualização (célula D3).
  Quando a obra tem **uma linha por torre** (ex.: Origem Itaquera, torres 05, 06 e
  07), elas viram **uma linha por obra pela média simples das torres** — cada torre
  pesa igual —, e as colunas de máxima ficam com o maior valor entre as torres.
  O painel calcula do mesmo jeito: o P15 da média tratada é feito **dentro de cada
  torre** e a obra é a média simples das torres (Origem Itaquera: 21,63 na média
  tratada). Cada aba `Prod_` é uma torre;
- abas com **`OBSOLET`** no nome (ex.: `Prod_Origem-Itaquera_T7_OBSOLET`) são
  **ignoradas**, para não contar a mesma obra duas vezes.

As colunas são procuradas **pelo nome do cabeçalho**, não pela posição. Por isso
as abas podem ter colunas diferentes entre si — `EMPREITEIRO` em umas e não em
outras, `TORRE` uma vez ou seis vezes — sem quebrar nada. Inserir ou mover coluna
também é seguro.

O que quebra:

- mudar o prefixo `Prod_` das abas de obra, ou renomear a aba `GERAL`;
- renomear cabeçalhos como `SOMA DA ÁREA DE FÔRMA MONTADA NA SEMANA (m2)`,
  `QTDE. MÉDIA DE MONTADORES NA SEMANA` ou `PRODUTIVIDADE MÉDIA DIÁRIA POR
  MONTADOR (m2/dia)`;
- apagar a coluna `OBRA`, que é a âncora usada para localizar o cabeçalho.

A linha do cabeçalho e as colunas do quadro `GERAL` são descobertas
automaticamente, então mover o cabeçalho de linha ou reordenar o quadro não
atrapalha.

Depois de cada alteração de estrutura, rode o workflow pelo botão **Run workflow**
e confira o log antes de confiar no painel.

Para fixar a cor de uma obra, edite o mapa `HEX` no `index.html`.

## Data de atualização do painel

Cabeçalho e rodapé mostram a constante `DATA_ATUALIZACAO_FIXA` do
`scripts/gerar_dados.py` (hoje `2026-10-19`). Para voltar a usar a data da planilha
(aba GERAL, `DATA ATUALIZAÇÃO`), deixe a constante como `None`.

Os filtros de período **Semanal, Mensal e Anual** abrem sempre no período mais
recente com produção apontada (sem passar da data de atualização).

## Gráfico 4 — Evolução (eixo "calendário")

O eixo acompanha o filtro de período: **Semanal** mostra semana a semana todas as
semanas do mês escolhido; **Mensal**, mês a mês todos os meses do ano escolhido;
**Anual**, ano a ano todos os anos com dados; **Personalizado** escolhe pela duração
(até 62 dias semana a semana, até cerca de 2 anos mês a mês, acima disso ano a ano);
**Todo Período** segue semana a semana. Em cada período a obra tem **um ponto só**:
cada torre entra com a sua média (simples, ou tratada pelo P15 da própria torre) e,
havendo mais de uma torre no mesmo período, vale o maior ponto.

