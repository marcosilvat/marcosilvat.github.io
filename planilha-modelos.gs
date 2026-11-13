/**
 * Recebe e entrega os modelos de mensagem do site (modelos.html).
 * Cole este código em Extensões → Apps Script da planilha e publique como App da Web.
 * A chave do time fica em Configurações do projeto → Propriedades do script → CHAVE.
 */
const ABA = 'Modelos';
const COLUNAS = ['id', 'criado_em', 'title', 'cat', 'text', 'vars', 'button', 'ref', 'pdf', 'pix', 'barcode'];
const GRUPOS = ['lembrete', 'cobranca', 'negociacao', 'pagamento', 'aviso', 'manutencao', 'contrato'];

function aba_() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(ABA);
  if (!aba) {
    aba = planilha.insertSheet(ABA);
    aba.appendRow(COLUNAS);
  }
  return aba;
}

function resposta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Evita que um texto começando com = + - @ vire fórmula na planilha
function texto_(v) {
  const t = String(v || '').slice(0, 1500);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function doGet() {
  const linhas = aba_().getDataRange().getValues();
  const cabecalho = linhas.shift();
  const modelos = linhas
    .filter(l => l[0])
    .map(l => Object.fromEntries(cabecalho.map((c, i) => [c, l[i]])));
  return resposta_({ ok: true, modelos });
}

function doPost(e) {
  let corpo;
  try {
    corpo = JSON.parse(e.postData.contents);
  } catch (err) {
    return resposta_({ ok: false, erro: 'Pedido inválido' });
  }

  const chave = PropertiesService.getScriptProperties().getProperty('CHAVE');
  if (!chave || corpo.chave !== chave) return resposta_({ ok: false, erro: 'Chave do time inválida' });

  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    const aba = aba_();

    if (corpo.acao === 'excluir') {
      const total = Math.max(aba.getLastRow() - 1, 0);
      const ids = total ? aba.getRange(2, 1, total, 1).getValues().map(l => String(l[0])) : [];
      const i = ids.indexOf(String(corpo.id));
      if (i < 0) return resposta_({ ok: false, erro: 'Modelo não encontrado' });
      aba.deleteRow(i + 2);
      return resposta_({ ok: true });
    }

    const m = corpo.modelo || {};
    if (!String(m.title || '').trim() || !String(m.text || '').trim()) {
      return resposta_({ ok: false, erro: 'Nome e texto são obrigatórios' });
    }
    const id = Utilities.getUuid();
    aba.appendRow([
      id, new Date(), texto_(m.title), GRUPOS.indexOf(m.cat) >= 0 ? m.cat : 'aviso',
      texto_(m.text), texto_(m.vars), texto_(m.button),
      m.ref === true, m.pdf === true, m.pix === true, m.barcode === true
    ]);
    return resposta_({ ok: true, id });
  } finally {
    trava.releaseLock();
  }
}
