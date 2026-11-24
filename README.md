# Leozin Barber Studio — site de agendamento

Site estático (GitHub Pages) com agendamento online ligado a uma planilha do Google por meio de um Web App do Google Apps Script.

> Todos os dados deste repositório são de exemplo e estão marcados como **[EXEMPLO]** ou **[a definir]**: endereços, telefones, barbeiros e Instagram. Nenhum dado real de clientes, e-mail ou chave fica aqui.

## Estrutura

| Caminho | O que é |
| --- | --- |
| `index.html` | O site inteiro: página, pop-up de agendamento e área "Meus agendamentos" |
| `assets/` | Logo e fotos |
| `apps-script/Code.gs` | Back-end (Google Apps Script) |
| `apps-script/appsscript.json` | Manifesto do Apps Script (fuso America/Sao_Paulo, Web App) |

## Como funciona

1. O cliente cria a conta (nome, e-mail, senha e aceite LGPD) ou entra.
2. Escolhe a unidade, o serviço, o barbeiro, o dia (próximos 7 dias) e um horário livre (blocos de 30 min).
3. Confirma e recebe um e-mail de confirmação. 30 minutos antes recebe um lembrete.
4. Em **Meus agendamentos** ele remarca ou cancela até **2 h antes** do horário. Cada ação envia um e-mail.

Regras atuais (todas editáveis na aba **Config** da planilha): terça a sábado, 9h às 19h, almoço 12h às 14h, domingo e segunda fechados.

## Instalação do back-end (uma vez, uns 5 minutos)

1. Logado na conta Google da barbearia, abra <https://sheets.new> e dê um nome à planilha, por exemplo "Leozin — Agenda".
2. Na planilha, vá em **Extensões → Apps Script**.
3. Apague o conteúdo de `Código.gs` e cole todo o conteúdo de [`apps-script/Code.gs`](apps-script/Code.gs). Salve.
4. Em **Configurações do projeto** (ícone de engrenagem), defina o fuso horário como **(GMT-03:00) São Paulo**.
5. No editor, selecione a função **`setup`** e clique em **Executar**. Autorize o acesso quando o Google pedir. A função:
   - cria as abas **Agenda**, **Clientes**, **Config** e **Bloqueios**, com os dados de exemplo;
   - cria o gatilho `enviarLembretes`, que roda a cada 5 minutos;
   - gera o segredo (`PEPPER`) usado nas senhas, guardado nas Propriedades do Script.
6. Clique em **Implantar → Nova implantação → Tipo: App da Web**.
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
7. Copie a URL que termina em `/exec`.
8. Em `index.html`, cole essa URL na linha `const API_URL = ...` (logo no começo do `<script>`) e publique.

**Opcional, em Configurações do projeto → Propriedades do script:**
- `SITE_URL`: endereço do site publicado. Ativa o botão "Meus agendamentos" nos e-mails.
- `REPLY_TO`: endereço de resposta dos e-mails. Se não existir, usa o e-mail da conta que instalou o script.

Sempre que alterar o `Code.gs`, faça **Implantar → Gerenciar implantações → Editar → Nova versão** para a URL `/exec` passar a usar o código novo.

## A planilha

| Aba | Conteúdo |
| --- | --- |
| **Agenda** | Um agendamento por linha: unidade, serviço, barbeiro, data, hora, cliente, status (`confirmado` / `cancelado`), lembrete enviado e histórico de mudanças |
| **Clientes** | Nome, e-mail, salt e hash da senha, data do aceite LGPD e hash do token de sessão. A senha nunca é gravada |
| **Config** | Parâmetros (horários, almoço, dias, prazo, lembrete), unidades, barbeiros e serviços. Para desativar algo, escreva `não` na coluna **Ativo** |
| **Bloqueios** | Folgas e ausências: Unidade, Barbeiro, Data (`aaaa-mm-dd`), Início, Fim (`HH:mm`) e Motivo. Deixe Unidade ou Barbeiro vazio para valer para todos, a Data vazia para todos os dias e Início/Fim vazios para o dia inteiro |

Os IDs das unidades, barbeiros e serviços (coluna **ID** da Config) são usados pelo site. Ao trocar nomes, mantenha os IDs. Se criar um barbeiro novo, a foto dele no pop-up vem do objeto `BARBERS` do `index.html` (sem foto, aparece a inicial do nome).

## Segurança e privacidade

- Senhas: SHA-256 com salt aleatório por cliente, um segredo extra (`PEPPER`) fora da planilha e 1.000 rodadas. O Apps Script não oferece bcrypt nem Argon2.
- Sessão: o navegador guarda um token aleatório; a planilha guarda só o hash dele. Vale por 30 dias e é trocado a cada login.
- Cada ação confere se o agendamento é do cliente logado. A consulta de horários não expõe nomes nem e-mails.
- 5 senhas erradas seguidas bloqueiam o e-mail por 15 minutos.
- Agendar e remarcar usam `LockService` e conferem de novo se o horário está livre dentro do bloqueio, para não haver dois clientes no mesmo horário.
- Quem tem acesso à planilha vê os dados dos clientes. Compartilhe só com quem precisa.
- **Política de Privacidade:** o link no site aponta para `#politica-de-privacidade`, que ainda não existe. Publique a política antes de abrir o agendamento ao público.

## CORS

O site chama o Web App com `POST` e `Content-Type: text/plain`, que não dispara a verificação prévia (preflight) do navegador, algo que o Apps Script não responde. O Google responde por um redirecionamento para `googleusercontent.com`, que já libera o acesso de qualquer origem. Nada precisa ser configurado.

## Limites do Google

- Uma conta Gmail gratuita envia cerca de **100 e-mails por dia** pelo Apps Script. Confirmação, lembrete, remarcação e cancelamento contam. Se a cota acabar, o agendamento é salvo mesmo assim e só o e-mail deixa de sair. Para mais volume, use uma conta Google Workspace (1.500/dia).
- O Web App leva de 1 a 3 segundos por resposta.

## Publicação (GitHub Pages)

Em **Settings → Pages**, escolha **Deploy from a branch**, branch `main`, pasta `/ (root)`, e salve. O site fica em `https://<usuario>.github.io/leozin-barber-studio/`.
