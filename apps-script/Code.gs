
const CONFIG = {
  SHEET_NAME: 'Ferias', 
  SLACK_TOKEN: '',


  COL_COLABORADOR: 1,   
  COL_DATA_INICIO: 2,   
  COL_DATA_FIM: 3,      
  COL_GESTOR: 4,        
  COL_SLACK_ID: 5,      
  COL_CIENTE: 6,        
  COL_OBSERVACAO: 7,    
  COL_NOTIFICAR: 8,     
  COL_ENVIADO: 9     
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏖️ Gestão de Férias')
    .addItem('📤 Enviar avisos para gestores', 'enviarAvisoFeriasGestor')
    .addSeparator()
    .addItem('ℹ️ Sobre', 'mostrarInfo')
    .addToUi();
}

function mostrarInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'ℹ️ Sistema de Gestão de Férias',
    '📤 Enviar avisos para gestores:\n' +
    'Envia notificações via Slack para os gestores sobre férias dos colaboradores.\n\n' +
    '✅ Somente linhas com "SIM" na coluna "Notificar" serão processadas.\n' +
    '⏭️ Linhas já marcadas como "Ciente" serão puladas.\n\n' +
    'Os gestores poderão:\n' +
    '• Marcar como ciente\n' +
    '• Adicionar observações',
    ui.ButtonSet.OK
  );
}

function enviarAvisoFeriasGestor() {
  const SLACK_URL = 'https://slack.com/api/chat.postMessage';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Erro: Planilha "' + CONFIG.SHEET_NAME + '" não encontrada!');
    return;
  }

  if (!CONFIG.SLACK_TOKEN.trim()) {
    SpreadsheetApp.getUi().alert('❌ Erro: Token do Slack não configurado!');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let enviados = 0;
  let erros = [];
  let pulados = 0;

  for (let i = 1; i < data.length; i++) {
    const colaborador = data[i][CONFIG.COL_COLABORADOR - 1];
    const rawInicio = data[i][CONFIG.COL_DATA_INICIO - 1];
    const rawFim = data[i][CONFIG.COL_DATA_FIM - 1];
    const gestor = data[i][CONFIG.COL_GESTOR - 1];
    const slackIdGestor = data[i][CONFIG.COL_SLACK_ID - 1];
    const ciente = data[i][CONFIG.COL_CIENTE - 1];
    const notificar = (data[i][CONFIG.COL_NOTIFICAR - 1] || '').toString().trim().toUpperCase();

    if (notificar !== 'SIM' || !colaborador || !rawInicio || !slackIdGestor || ciente) {
      pulados++;
      continue;
    }

    try {
      const dataInicio = new Date(rawInicio);
      const dataInicioStr = Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy');

      const dataFim = rawFim ? new Date(rawFim) : null;
      const dataFimStr = dataFim ? Utilities.formatDate(dataFim, Session.getScriptTimeZone(), 'dd/MM/yyyy') : null;

      const payload = {
        channel: slackIdGestor.toString().trim(),
        text: `Aviso de férias: ${colaborador}`,
        blocks: [
          {
            type: "section",
            text: {
              type: "mrkdwn",
              text: `👋 Olá *${gestor || 'Gestor'}*,\n\nO colaborador *${colaborador}* inicia férias em *${dataInicioStr}*` + (dataFimStr ? ` até *${dataFimStr}*.` : '.') + `\n\nPor favor, selecione uma das opções abaixo.`
            }
          },
          {
            type: "actions",
            elements: [
              {
                type: "button",
                text: { 
                  type: "plain_text", 
                  text: "✅ Ciente", 
                  emoji: true 
                },
                style: "primary",
                action_id: "ferias_ciente",
                value: `${i + 1}|${colaborador}`
              },
              {
                type: "button",
                text: { 
                  type: "plain_text", 
                  text: "📝 Adicionar observação", 
                  emoji: true 
                },
                action_id: "ferias_obs",
                value: `${i + 1}|${colaborador}`
              }
            ]
          }
        ]
      };

      Logger.log(`Enviando para ${slackIdGestor}: ${JSON.stringify(payload)}`);

      const response = UrlFetchApp.fetch(SLACK_URL, {
        method: 'post',
        contentType: 'application/json',
        headers: { 
          'Authorization': `Bearer ${CONFIG.SLACK_TOKEN.trim()}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const result = JSON.parse(response.getContentText());
      
      Logger.log(`Resposta para ${colaborador}: ${JSON.stringify(result)}`);

      if (result.ok) {
        sheet.getRange(i + 1, CONFIG.COL_ENVIADO).setValue(
          `✅ Enviado em ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}`
        );
        enviados++;
      } else {
        erros.push(`${colaborador}: ${result.error || 'Erro desconhecido'}`);
        Logger.log(`Erro detalhado: ${JSON.stringify(result)}`);
      }

      Utilities.sleep(1000);

    } catch (error) {
      erros.push(`${colaborador}: ${error.message}`);
      Logger.log(`Exceção para ${colaborador}: ${error.message}`);
    }
  }

  let msg = `✅ Processo concluído!\n\n📤 Enviados: ${enviados}\n⏭️ Pulados: ${pulados}`;
  if (erros.length > 0) msg += `\n\n❌ Erros:\n${erros.join('\n')}`;

  SpreadsheetApp.getUi().alert(msg);
}

function salvarObservacao(row, observacao) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  sheet.getRange(row, CONFIG.COL_OBSERVACAO).setValue(observacao);
}

function abrirModal(triggerId, row, colaborador) {
  const modal = {
    trigger_id: triggerId,
    view: {
      type: "modal",
      callback_id: "salvar_observacao",
      title: { type: "plain_text", text: "Observação", emoji: true },
      submit: { type: "plain_text", text: "Salvar", emoji: true },
      close: { type: "plain_text", text: "Cancelar", emoji: true },
      private_metadata: `${row}|${colaborador}`,
      blocks: [
        {
          type: "input",
          block_id: "obs_block",
          element: {
            type: "plain_text_input",
            multiline: true,
            action_id: "obs_text",
            placeholder: { type: "plain_text", text: "Digite sua observação aqui..." }
          },
          label: { type: "plain_text", text: `Observação sobre ${colaborador}:` }
        }
      ]
    }
  };

  UrlFetchApp.fetch("https://slack.com/api/views.open", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(modal),
    headers: { Authorization: `Bearer ${CONFIG.SLACK_TOKEN}` },
    muteHttpExceptions: true
  });
}

function doPost(e) {
  try {
    if (!e || !e.parameter) return retornoSucesso();
    const payloadStr = e.parameter.payload;
    if (!payloadStr) return retornoSucesso();

    const payload = JSON.parse(payloadStr);

    if (payload.actions && payload.actions[0]) {
      const action = payload.actions[0];

      if (action.action_id === "ferias_obs") {
        const [row, colaborador] = action.value.split('|');
        const props = PropertiesService.getScriptProperties();
        props.setProperty('response_url_' + row, payload.response_url);
        abrirModal(payload.trigger_id, parseInt(row), colaborador);
        return retornoSucesso();
      }

      if (action.action_id === "ferias_ciente") {
        const [row, colaborador] = action.value.split('|');
        marcarComoCiente(parseInt(row));
        atualizarMensagem(payload.response_url, colaborador, 'ciente');
        return retornoSucesso();
      }
    }

    if (payload.type === "view_submission" && payload.view.callback_id === "salvar_observacao") {
      const [row, colaborador] = payload.view.private_metadata.split('|');
      const observacao = payload.view.state.values.obs_block.obs_text.value;
      const userId = payload.user.id;
      processarObservacaoBackground(parseInt(row), colaborador, observacao, userId);
      return ContentService.createTextOutput(JSON.stringify({ response_action: "clear" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return retornoSucesso();

  } catch (error) {
    Logger.log(`Erro no doPost: ${error.message}`);
    return retornoSucesso();
  }
}

function retornoSucesso() {
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function processarObservacaoBackground(row, colaborador, observacao, userId) {
  try {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);
    salvarObservacao(row, observacao);
    const props = PropertiesService.getScriptProperties();
    const responseUrl = props.getProperty('response_url_' + row);
    if (responseUrl) {
      atualizarMensagem(responseUrl, colaborador, 'observacao');
      props.deleteProperty('response_url_' + row);
    }
    enviarMensagemConfirmacao(userId, colaborador, observacao);
    lock.releaseLock();
  } catch (e) {
    Logger.log('Erro no background: ' + e.message);
  }
}

function atualizarMensagem(responseUrl, colaborador, tipo) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const mensagem = {
    replace_original: true,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text:
            tipo === 'ciente'
              ? `✅ *Confirmado!*\n\nVocê marcou como ciente as férias de *${colaborador}*.\n\n_Registrado em: ${now}_`
              : `✅ *Observação registrada!*\n\nSua observação sobre as férias de *${colaborador}* foi salva com sucesso.\n\n_Registrado em: ${now}_`
        }
      }
    ]
  };

  UrlFetchApp.fetch(responseUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(mensagem),
    muteHttpExceptions: true
  });
}

function enviarMensagemConfirmacao(userId, colaborador, observacao) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const payload = {
    channel: userId,
    text: `Observação registrada para ${colaborador}`,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text:
            `✅ *Observação registrada com sucesso!*\n\n*Colaborador:* ${colaborador}\n*Sua observação:* ${observacao}\n\n_Registrado em: ${now}_`
        }
      }
    ]
  };

  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${CONFIG.SLACK_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function marcarComoCiente(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  sheet.getRange(row, CONFIG.COL_CIENTE).setValue(`✅ Ciente em ${now}`);
}
