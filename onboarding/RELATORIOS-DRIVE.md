# Receber os resultados das provas no seu Google Drive

A plataforma já registra, por pessoa, o resultado das provas e os **pontos de
atenção** (temas onde o fechador errou). Para que esses resultados cheguem
**automaticamente ao seu Drive**, faça este passo a passo **uma única vez**:

## 1. Crie a planilha que vai receber os resultados
1. Acesse https://drive.google.com → **Novo → Planilhas Google**.
2. Renomeie para **"Resultados Onboarding"**.

## 2. Cole o script
1. Na planilha: menu **Extensões → Apps Script**.
2. Apague o conteúdo e cole o código abaixo. Salve (ícone de disquete).

```javascript
function doPost(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Resultados') || ss.insertSheet('Resultados');
  var d = JSON.parse(e.postData.contents);
  if (sheet.getLastRow() === 0){
    sheet.appendRow(['Data','Nome','E-mail','Operacao','Nota geral %','Acertos','Total',
                     'Assuntos com dificuldade','Detalhe dos pontos de atencao','JSON']);
  }
  var pa = d.pontos_de_atencao || [];
  var assuntos = pa.map(function(x){return x.assunto;});
  var uniq = assuntos.filter(function(v,i){return assuntos.indexOf(v)===i;});
  var detalhe = pa.map(function(x){
    return '- ['+x.assunto+'] '+x.questao+'\n   respondeu: '+x.respondeu+'\n   esperado: '+x.esperado;
  }).join('\n');
  sheet.appendRow([new Date(), d.aluno.nome, d.aluno.email, d.aluno.operacao,
                   d.nota_geral_pct, d.acertos, d.total, uniq.join('; '), detalhe, JSON.stringify(d)]);
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
                       .setMimeType(ContentService.MimeType.JSON);
}
```

## 3. Publique como aplicativo web
1. Botão **Implantar → Nova implantação**.
2. Em "Tipo", escolha **App da Web**.
3. "Executar como": **Eu**. "Quem pode acessar": **Qualquer pessoa**.
4. **Implantar** e autorize o acesso (é a sua própria conta).
5. Copie a **URL do app da Web** (termina em `/exec`).

## 4. Ligue a plataforma à sua planilha
- Me envie essa URL **ou** abra o arquivo `onboarding/index.html`, procure a linha
  `const RELATORIO_ENDPOINT='';` e cole a URL entre as aspas. Ex.:
  `const RELATORIO_ENDPOINT='https://script.google.com/macros/s/AAAA.../exec';`

Pronto. A partir daí, toda vez que um fechador clicar em **"Enviar resultado para
a gestão"** na aba *Minha avaliação*, uma nova linha aparece na sua planilha, já
com os **assuntos de maior dificuldade** destacados.

> Enquanto a URL não estiver configurada, o botão continua funcionando: ele **baixa
> um relatório** (HTML) e abre um **e-mail já preenchido** para você — é só anexar o
> arquivo e enviar.

Observação: o acesso "Qualquer pessoa" vale só para o endereço secreto do seu
app; ele apenas **recebe** os resultados do treinamento e grava na sua planilha.
