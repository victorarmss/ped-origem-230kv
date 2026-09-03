# Acompanhamento Geral de Projetos — Publicação Automática (GitHub Pages)

Dashboard de acompanhamento dos projetos **SE UTE Pilar 230 kV**, **SE Rio Largo II** e **Linha de Transmissão**.

Este guia configura a **atualização automática**: quando você **salva a planilha `PED - ORIGEM 230kV.xlsx`**, o dashboard é regenerado e **publicado sozinho** na web, usando o **GitHub** (hospedagem via **GitHub Pages** — sem Netlify).

---

## Como funciona

```
Você salva a PED  ──►  Vigia detecta  ──►  regenera o index.html  ──►  git push  ──►  GitHub Pages publica
```

- **`index.html`** é o dashboard (com os dados da PED embutidos).
- O **Vigia** (`Vigiar_PED.ps1`) fica rodando na sua máquina monitorando a PED.
- Ao salvar a PED, ele chama o **`Publicar_Auto.ps1`**, que regenera o `index.html` e faz `git push`.
- O **GitHub Pages** publica a nova versão em ~1–2 minutos.

> ⚠️ Como os dados vêm de um **Excel na sua máquina**, o vigia precisa estar **rodando no seu computador**. Não é 100% nuvem.

---

## Pré-requisitos

- **Git** — já instalado nesta máquina ✅ (versão verificada: 2.55).
- Uma conta no **GitHub** — https://github.com (grátis).

---

## Passo 1 — Criar o repositório no GitHub

1. Entre em https://github.com/new
2. **Repository name:** `ped-origem-230kv` (ou o nome que preferir).
3. **Visibilidade:** marque **Public**.
   > No plano **gratuito**, o GitHub Pages só funciona com repositório **Public**. (Repositório privado com Pages exige o plano pago GitHub Pro.) Veja a nota de **Privacidade** no fim.
4. **Não** marque "Add a README" (a pasta já tem os arquivos).
5. Clique **Create repository** e **copie a URL**, algo como
   `https://github.com/SEU-USUARIO/ped-origem-230kv.git`

## Passo 2 — Enviar esta pasta para o GitHub

Abra o **PowerShell** nesta pasta (Shift + botão direito na pasta → "Abrir janela do PowerShell aqui") e rode, **trocando a URL** pela sua:

```powershell
cd "$env:USERPROFILE\OneDrive - ENGETECNICA\Área de Trabalho\PED Claude"
git init
git branch -M main
git add .
git commit -m "Primeira versao do dashboard"
git remote add origin https://github.com/SEU-USUARIO/ped-origem-230kv.git
git push -u origin main
```

No primeiro `push`, o Git **abre o navegador para você entrar no GitHub** — faça o login e autorize. Os arquivos sobem (a planilha PED **não** sobe, está no `.gitignore`).

## Passo 3 — Ligar o GitHub Pages

1. No repositório, vá em **Settings** (aba no topo) → **Pages** (menu à esquerda).
2. Em **Build and deployment → Source**, escolha **Deploy from a branch**.
3. Em **Branch**, selecione **`main`** e a pasta **`/ (root)`**, e clique **Save**.
4. Aguarde ~1–2 min. A URL do dashboard aparece no topo dessa mesma tela, algo como:
   `https://SEU-USUARIO.github.io/ped-origem-230kv/`

> A partir daí, **todo `git push` republica automaticamente** nessa URL.

## Passo 4 — Ativar a publicação automática ao salvar a PED

Dê **dois cliques** em **`Iniciar Vigia.cmd`**. Abre uma janelinha (pode minimizar) que fica vigiando a PED.

- **Teste:** abra a PED, faça uma alteração, **salve e feche**. Em ~15 s o vigia dispara o `git push`; o GitHub Pages publica em seguida. Acompanhe pelo arquivo **`publicar_log.txt`**.
- Para **parar**, feche a janela do vigia.

## Passo 5 (opcional) — Iniciar o vigia junto com o Windows

1. Aperte **Win + R**, digite `shell:startup` e Enter.
2. Coloque nessa pasta um **atalho** do `Iniciar Vigia.cmd` (botão direito no arquivo → *Enviar para → Área de trabalho (criar atalho)*, depois mova o atalho para a pasta que abriu).

Assim o vigia sobe sozinho quando você liga o computador.

---

## Uso no dia a dia

Depois de configurado, é só **trabalhar na PED normalmente**. Cada vez que salvar, o dashboard republica sozinho. Nada mais a fazer.

## Publicar manualmente (sem esperar o vigia)

- Botão direito em **`Publicar_Auto.ps1`** → *Executar com PowerShell* — regenera e faz o push na hora.
- Ou **`Atualizar Dashboard.cmd`** — regenera o `index.html` e **abre no navegador** (não publica; use só para conferir localmente).

## Solução de problemas

- **Nada publicou:** confira o `publicar_log.txt`. Se disser "ERRO no git push", rode um `git push` manual no PowerShell para reautenticar no GitHub.
- **A página não abre / dá 404:** confirme o Passo 3 (Pages ligado em `main` / root) e espere 1–2 min após o push.
- **"a pasta ainda nao e um repositorio git":** faça o Passo 2.
- **O vigia não dispara:** confirme que a janela do `Iniciar Vigia.cmd` está aberta e que você **salvou** a PED (Ctrl + S).
- **Mudou as colunas da PED?** O leitor usa posições fixas (1ª emissão = coluna 14, status = coluna 13, reprogramado = coluna 10 "DATA PLANEJADA REVISÃO"). Se reorganizar a planilha, me avise para reajustar.

## Privacidade — leia isto

- Com o **GitHub Pages gratuito**, o repositório é **Public** → o `index.html` (que contém os dados dos projetos) fica **acessível publicamente** (tanto no repositório quanto na URL `github.io`).
- A **planilha PED não é enviada** ao GitHub (está no `.gitignore`), mas os **números/documentos** que aparecem no dashboard ficam públicos.
- Se os dados **não podem ser públicos**, as opções são: (a) **GitHub Pro** (pago) para repositório privado com Pages; ou (b) voltar a um hospedeiro com proteção por senha. Me avise que eu ajusto.

## Arquivos desta pasta

| Arquivo | Função |
|---|---|
| `PED - ORIGEM 230kV.xlsx` | A planilha-fonte (fica só na sua máquina) |
| `index.html` | O dashboard publicado |
| `Atualizar_Dashboard.ps1` | Regenera o `index.html` a partir da PED |
| `Publicar_Auto.ps1` | Regenera + envia ao GitHub (Pages publica) |
| `Vigiar_PED.ps1` | Monitora a PED e dispara a publicação ao salvar |
| `Iniciar Vigia.cmd` | Liga o vigia |
| `Atualizar Dashboard.cmd` | Atualiza e abre o dashboard localmente (manual) |
| `Logo.png` | Logo da Engetécnica |
| `.gitignore`, `.nojekyll` | Config do Git / do GitHub Pages |
| `publicar_log.txt` | Histórico das publicações (gerado automaticamente) |
