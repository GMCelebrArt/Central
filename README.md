# GM CelebrArt — Central

Página: **GitHub Pages** (`index.html`). Dados: **Google Planilha + Drive**, via Apps Script (pasta `apps-script/`), protegidos por PIN.

## Publicar a página
Settings → Pages → Source: *Deploy from a branch* → Branch `main` / `(root)` → Save.
Endereço: `https://gmcelebrart.github.io/Central/`

## Apps Script (servidor)
1. Cole `apps-script/Code.gs` no editor do Apps Script da planilha.
2. Implantar → Gerenciar implantações → ✏️ → **Nova versão** → Implantar
   (Executar como: **Eu** · Quem pode acessar: **Qualquer pessoa**).
3. Se a URL do App da Web mudar, atualize `API_URL` no `index.html`.

## PIN
Aba **Config** da planilha, linha `pin`. Trocar lá vale na hora; cada aparelho pede o novo PIN.
8 PINs errados bloqueiam o acesso por 10 minutos.

## Editar dados
Acervo, kits, peças do checklist, extras, temas e dados da empresa: direto na planilha.
