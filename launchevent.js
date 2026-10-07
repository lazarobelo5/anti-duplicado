/*
 * Anti-Duplicado AL-COX — bloqueia o envio de e-mail NOVO com título já usado.
 * Roda no New Outlook / Outlook Web (Smart Alerts, evento OnMessageSend).
 *
 * Regras:
 *  - Respostas e encaminhamentos (RE:, RES:, ENC:, FW:, FWD:, TR:) passam sempre.
 *  - Título comparado sem acento, sem diferença de maiúscula/minúscula e espaços extras.
 *  - O histórico fica nas configurações da própria caixa (roamingSettings).
 */

// ===== CONFIGURAÇÃO =====
var CHAVE = "alcox_titulos_v1";
var MAX_TITULOS = 180;          // limite de armazenamento do Outlook (~32 KB)
// Deixe null para verificar TODOS os e-mails novos.
// Exemplo para verificar só os do processo: /AL-COX|4501\d{6}|PEDIDO DE COMPRA/i
var FILTRO_TITULO = null;
// ========================

var PREFIXO_RESPOSTA = /^\s*((re|res|enc|fw|fwd|tr)\s*:\s*)+/i;

function normalizar(titulo) {
  return (titulo || "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function ehRespostaOuEncaminhamento(titulo) {
  return PREFIXO_RESPOSTA.test(titulo || "");
}

function hojeBR() {
  var d = new Date();
  function p(n) { return (n < 10 ? "0" : "") + n; }
  return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear() +
    " " + p(d.getHours()) + ":" + p(d.getMinutes());
}

function lerHistorico() {
  var lista = Office.context.roamingSettings.get(CHAVE);
  return Array.isArray(lista) ? lista : [];
}

function onMessageSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync(function (res) {
    // Se não conseguir ler o título, não trava o trabalho.
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    var titulo = res.value || "";

    if (!titulo.trim()) {
      event.completed({ allowEvent: false, errorMessage: "O e-mail está sem título. Preencha o assunto antes de enviar." });
      return;
    }

    if (ehRespostaOuEncaminhamento(titulo) || (FILTRO_TITULO && !FILTRO_TITULO.test(titulo))) {
      event.completed({ allowEvent: true });
      return;
    }

    var chave = normalizar(titulo);
    var lista = lerHistorico();
    var achado = null;
    for (var i = 0; i < lista.length; i++) {
      if (lista[i].k === chave) { achado = lista[i]; break; }
    }

    if (achado) {
      event.completed({
        allowEvent: false,
        errorMessage: "TÍTULO REPETIDO: já foi enviado um e-mail com o assunto \"" + titulo.trim() +
          "\" em " + (achado.d || "data não registrada") +
          ". Altere o título (ex.: novo sequencial) ou responda o e-mail original."
      });
      return;
    }

    lista.unshift({ k: chave, d: hojeBR() });
    if (lista.length > MAX_TITULOS) lista = lista.slice(0, MAX_TITULOS);
    Office.context.roamingSettings.set(CHAVE, lista);
    Office.context.roamingSettings.saveAsync(function () {
      event.completed({ allowEvent: true });
    });
  });
}

// Registro do evento (obrigatório para New Outlook, Outlook Web e Outlook clássico)
if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

// Exportação para teste local (Node) — ignorada pelo Outlook
if (typeof module !== "undefined") {
  module.exports = { normalizar: normalizar, ehRespostaOuEncaminhamento: ehRespostaOuEncaminhamento };
}
