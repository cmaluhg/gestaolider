# Como configurar o backend Google Drive

Siga os passos abaixo **uma única vez**. Depois, o botão "Salvar no Drive" funcionará automaticamente.

---

## Passo 1 — Criar a pasta no Google Drive

1. Acesse [drive.google.com](https://drive.google.com)
2. Clique em **+ Novo → Pasta**
3. Nomeie como **`Relatórios — Painel da Líder`**
4. Abra a pasta e copie o ID da URL:
   - URL exemplo: `https://drive.google.com/drive/folders/`**`1aBcDeFgHiJkLmNoPqRsTuVwXyZ`**
   - O ID é a parte em negrito acima

---

## Passo 2 — Criar a planilha no Google Sheets

1. Acesse [sheets.google.com](https://sheets.google.com)
2. Crie uma planilha em branco
3. Nomeie como **`Relatórios — Painel da Líder`**
4. Copie o ID da URL:
   - URL exemplo: `https://docs.google.com/spreadsheets/d/`**`1ZyXwVuTsRqPoNmLkJiHgFeDcBa`**`/edit`
   - O ID é a parte em negrito

---

## Passo 3 — Criar o Google Apps Script

1. Acesse [script.google.com](https://script.google.com)
2. Clique em **+ Novo projeto**
3. Nomeie o projeto como **`Painel da Líder — Backend`**
4. **Apague** o código padrão (`function myFunction() {}`)
5. **Cole** todo o conteúdo do arquivo `Code.gs` deste repositório
6. No início do código, preencha:
   ```javascript
   var SHEET_ID  = 'SEU_ID_DA_PLANILHA';   // passo 2
   var FOLDER_ID = 'SEU_ID_DA_PASTA';       // passo 1
   ```
7. Clique em **💾 Salvar**

---

## Passo 4 — Publicar como Web App

1. Clique em **Implantar → Nova implantação**
2. Em "Tipo", selecione **Aplicativo da Web**
3. Configure:
   - **Descrição:** `Painel da Líder v1`
   - **Executar como:** `Eu (seu e-mail)`
   - **Quem tem acesso:** `Qualquer pessoa`
4. Clique em **Implantar**
5. Autorize as permissões quando solicitado (Google vai pedir acesso ao Drive, Sheets e Docs)
6. **Copie a URL** exibida — será algo como:
   `https://script.google.com/macros/s/AKfycbx.../exec`

---

## Passo 5 — Configurar na plataforma

1. Abra o Painel de Controle da Líder no navegador
2. Vá para a aba **Relatório**
3. Expanda a seção **☁️ Integração Google Drive**
4. Cole a URL copiada no passo 4
5. Clique em **Salvar URL**

Pronto! A partir de agora, ao clicar em **☁️ Salvar no Drive**, o relatório será:
- ✅ Registrado na planilha com ID único (ex: `REL-2025-0001`)
- ✅ Criado como Google Doc na pasta configurada
- ✅ Link do Doc exibido diretamente no painel

---

## Solução de problemas

| Problema | Solução |
|----------|---------|
| "Não foi possível conectar ao Drive" | Verifique se a URL está correta e se o Web App está publicado como "Qualquer pessoa" |
| "Erro do servidor" | Abra o Apps Script → Execuções e veja o log de erro |
| Relatório não aparece na planilha | Confirme que o `SHEET_ID` está correto e que você tem permissão de edição |
| Precisa atualizar o código | No Apps Script, edite o código → Implantar → Gerenciar implantações → Editar → Nova versão |
