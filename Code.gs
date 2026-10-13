/**
 * Quiz +18 — recebe respostas e arquivos do site.
 * Grava na planilha, envia e-mail para você e guarda arquivos no Google Drive.
 *
 * COMO USAR
 * 1. Abra a planilha > Extensões > Apps Script e cole este código (substitui o antigo).
 * 2. Troque EMAIL abaixo pelo seu e-mail.
 * 3. Selecione a função "testarEmail" e clique em Executar. Aceite as permissões.
 * 4. Implantar > Gerenciar implantações > lápis (editar) > Versão: "Nova versão" > Implantar.
 *    Assim a URL continua a mesma e o index.html não precisa mudar.
 *    (Tipo: App da Web | Executar como: Eu | Quem pode acessar: Qualquer pessoa)
 */
const CONFIG = {
  EMAIL: "SEU_EMAIL@gmail.com",   // <- coloque aqui o e-mail que vai receber tudo
  ABA: "Respostas",               // nome da aba da planilha
  PASTA_ARQUIVOS: "Quiz - Arquivos recebidos"
};

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.tipo === "arquivo") salvarArquivo(d);
    else salvarRespostas(d);
    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}

function doGet() {
  return resposta({ ok: true, msg: "Quiz ativo" });
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function salvarRespostas(d) {
  delete d.tipo;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.ABA) || ss.insertSheet(CONFIG.ABA);

  // Mantém o cabeçalho em dia: cria colunas novas se aparecer pergunta nova
  let header = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const novas = Object.keys(d).filter(k => header.indexOf(k) === -1);
  if (novas.length) {
    header = header.concat(novas);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  sh.appendRow(header.map(h => (d[h] !== undefined ? d[h] : "")));

  enviarEmailRespostas(d);
}

function enviarEmailRespostas(d) {
  const apelido = d["Apelido"] || "Sem apelido";
  const dados = ["Apelido", "Modo", "Idade", "Gênero", "Altura (cm)", "Peso (kg)", "Estimativa mesmo gênero (%)"];
  let html = "<h2>Novo quiz respondido</h2><p><b>" + esc(apelido) + "</b></p><ul>";
  dados.forEach(k => { html += "<li><b>" + k + ":</b> " + esc(d[k]) + "</li>"; });
  html += "</ul><h3>Respostas</h3><ol>";
  Object.keys(d).filter(k => k.indexOf("Q ") === 0).forEach(k => {
    html += "<li>" + esc(k.replace(/^Q \d+ - /, "")) + "<br><i>" + esc(d[k] || "(sem resposta)") + "</i></li>";
  });
  html += "</ol>";
  MailApp.sendEmail({
    to: CONFIG.EMAIL,
    subject: "Quiz respondido por " + apelido,
    htmlBody: html
  });
}

function salvarArquivo(d) {
  const apelido = d["Apelido"] || "sem-apelido";
  const nome = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm") + "_" + apelido + "_" + d.nome;
  const blob = Utilities.newBlob(Utilities.base64Decode(d.base64), d.mime, nome);

  const it = DriveApp.getFoldersByName(CONFIG.PASTA_ARQUIVOS);
  const pasta = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.PASTA_ARQUIVOS);
  const arquivo = pasta.createFile(blob);

  MailApp.sendEmail({
    to: CONFIG.EMAIL,
    subject: "Arquivo recebido de " + apelido,
    htmlBody: "<p><b>" + esc(apelido) + "</b> enviou um arquivo.</p><p><a href='" + arquivo.getUrl() + "'>Abrir no Google Drive</a></p>",
    attachments: [blob]
  });
}

function esc(s) {
  return String(s === undefined || s === null ? "" : s).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Rode esta função uma vez para autorizar e conferir se o e-mail chega.
function testarEmail() {
  MailApp.sendEmail(CONFIG.EMAIL, "Teste do Quiz", "Se você recebeu isto, o envio por e-mail está funcionando.");
  DriveApp.getRootFolder(); // força a permissão do Drive
}
