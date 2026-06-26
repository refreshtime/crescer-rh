// ═══════════════════════════════════════════════════════════
//  CRESCER — ESTEIRA DIGITAL DE RH
//  Codigo.gs — Backend Google Apps Script (ES5 / Rhino)
//
//  Desenvolvido por DOMANI CONSULTORIA
//  Todos os direitos reservados © 2025 Domani Consultoria
//  Sistema desenvolvido sob encomenda para Crescer Consultoria
// ═══════════════════════════════════════════════════════════

var CONFIG = {
  SPREADSHEET_ID:   '1IMRXFrSpmn9JE8YNJQ4Ir_Qbz__y-MTW8UhTVvCyp80',
  DRIVE_ROOT_ID:    '1e8vpi06qcowiPZo5ctWQptxMo2Tb2aFi',
  DOCS_TEMPLATE_ID: '1D6TwHyWgA8sbG4n0-RGTax7jms3mBnKr0aw6qa5dN40',
  CALENDAR_ID:      'adminstrativocrescer@gmail.com'
};

// ───────────────────────────────────────────────────────────
//  doGet — leitura via GET
// ───────────────────────────────────────────────────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var result;
  if (action === 'getEmpresas') {
    result = getEmpresas();
  } else if (action === 'getColaboradores') {
    result = getColaboradores(e.parameter.empresa || '');
  } else if (action === 'getAtividades') {
    result = getAtividades();
  } else if (action === 'getLembretes') {
    result = getLembretes();
  } else {
    result = { ok: true, msg: 'CRESCER RH API' };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────────────────────
//  doPost — escrita via POST
// ───────────────────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    var dados   = payload.dados;

    if (action === 'processAdmissao')            result = processAdmissao(dados);
    else if (action === 'processAtestado')        result = processAtestado(dados);
    else if (action === 'processDesligamento')    result = processDesligamento(dados);
    else if (action === 'processCadastroCliente') result = processCadastroCliente(dados);
    else if (action === 'processVaga')            result = processVaga(dados);
    else if (action === 'processDiagnostico')         result = processDiagnostico(dados);
    else if (action === 'processDocumentosColaborador') result = processDocumentosColaborador(dados);
    else if (action === 'criarDocJornada')               result = criarDocJornada(dados);
    else if (action === 'gerarDocumento')                result = gerarDocumento(dados);
    else if (action === 'gerarPacote')                      result = gerarPacote(dados);
    else if (action === 'gerarPacoteDes')                   result = gerarPacoteDes(dados);
    else if (action === 'salvarAtividade')                  result = salvarAtividade(dados);
    else if (action === 'salvarEntrevistaDesligamento')      result = salvarEntrevistaDesligamento(dados);
    else if (action === 'salvarLembrete')                    result = salvarLembrete(dados);
    else if (action === 'concluirLembrete')                  result = concluirLembrete(dados);
    else result = { ok: false, erro: 'Acao desconhecida: ' + action };

  } catch (err) {
    result = { ok: false, erro: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────────────────────
//  getEmpresas
// ───────────────────────────────────────────────────────────
function getEmpresas() {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Base_Clientes');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) result.push({ razaoSocial: data[i][1], nomeFantasia: data[i][12] || data[i][1], cnpj: data[i][2] });
    }
    return result;
  } catch (err) {
    return [];
  }
}

// ───────────────────────────────────────────────────────────
//  processAdmissao
// ───────────────────────────────────────────────────────────
function processAdmissao(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Admissoes', [
      'Data/Hora', 'Nome', 'CPF', 'CTPS', 'Nascimento', 'Telefone', 'Empresa', 'Cargo',
      'Data Admissao', 'Salario', 'Tipo', 'Aceite LGPD',
      'ASO', 'Docs Contab', 'Ponto', 'VR', 'VT', 'Odonto', 'Saude', 'Conta', 'Termos Imp', 'Arquivado',
      'Observacoes', 'Link Pasta Drive', 'Link Doc Fisico'
    ]);

    var agora    = new Date();
    var dataStr  = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var linkPasta = '';
    var linkDoc   = '';
    var et        = dados.etapas || {};

    if (dados.tipo === 'digital') {
      var pastaEmpresa = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa);
      var pastaFunc    = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome);
      var pastaAdmDig  = _getOrCreateFolder(pastaFunc.getId(), 'Admissão');
      if (dados.rgBase64)  _salvarArquivo(pastaAdmDig, 'RG_' + dados.nome + '.pdf',  dados.rgBase64,  dados.rgMime);
      if (dados.cpfBase64) _salvarArquivo(pastaAdmDig, 'CPF_' + dados.nome + '.pdf', dados.cpfBase64, dados.cpfMime);
      linkPasta = pastaAdmDig.getUrl();
    } else {
      linkDoc = _gerarDocFisico(dados, agora);
    }

    sheet.appendRow([
      dataStr, dados.nome, dados.cpf, dados.ctps || '', dados.nascimento, dados.telefone,
      dados.empresa, dados.cargo || '', dados.dataAdm || '', dados.salario || '',
      dados.tipo === 'digital' ? 'Digital' : 'Fisico',
      dados.tipo === 'digital' ? (dados.aceite ? 'SIM' : 'NAO') : '-',
      et.aso || '', et.docsContab || '', et.ponto || '', et.vr || '',
      et.vt || '', et.odonto || '', et.saude || '', et.conta || '',
      et.termosImp || '', et.arquivado || '',
      dados.obs || '', linkPasta, linkDoc
    ]);

    // — Base_Colaboradores (base para Gerador de Documentos)
    var sheetColabs = _getOrCreateSheet(ss, 'Base_Colaboradores', [
      'Data Cadastro', 'Empresa', 'Nome', 'CPF', 'CTPS', 'Nascimento', 'Telefone', 'Cargo', 'Data Admissão', 'Status'
    ]);
    sheetColabs.appendRow([
      dataStr, dados.empresa, dados.nome, dados.cpf, dados.ctps || '',
      dados.nascimento || '', dados.telefone || '', dados.cargo || '', dados.dataAdm || '', 'Ativo'
    ]);

    // — Evento de admissão no calendário
    var dataAdm    = dados.dataAdm ? new Date(dados.dataAdm + 'T09:00:00') : agora;
    var dataAdmFim = new Date(dataAdm.getTime() + 3600000);
    _criarEventoCalendar(
      '✅ Admissão: ' + dados.nome + ' — ' + dados.empresa,
      dataAdm, dataAdmFim,
      'Cargo: ' + (dados.cargo || '') + ' | CPF: ' + (dados.cpf || '') + ' | Tipo: ' + (dados.tipo || '')
    );

    // — Tarefas pendentes no calendário para etapas não concluídas
    // Prazos calculados com base na data de admissão (CLT + boas práticas de RH)
    // Negativos = dias ANTES da admissão | Positivos = dias DEPOIS
    var etapaItens = [
      ['asoAgendado', 'ASO agendado e confirmado',           -3],  // 3 dias antes
      ['aso',         'Resultado do ASO: APTO',              -1],  // 1 dia antes
      ['ponto',       'Cadastrado no sistema de ponto',      -1],  // 1 dia antes
      ['docsContab',  'Documentação enviada à contabilidade', 2],  // +2 dias
      ['vr',          'Cadastrado no sistema VR',             3],  // +3 dias
      ['vt',          'Cadastrado no sistema de transporte',  3],  // +3 dias
      ['termosImp',   'Documentos assinados via ZipSign',     3],  // +3 dias
      ['conta',       'Abertura de conta solicitada',         5],  // +5 dias
      ['arquivado',   'Documentos digitalizados e arquivados',7],  // +7 dias
      ['odonto',      'Incluído no plano odontológico',      30],  // +30 dias (prazo legal)
      ['saude',       'Incluído no plano de saúde',          30]   // +30 dias (prazo legal)
    ];
    for (var ei = 0; ei < etapaItens.length; ei++) {
      var chave    = etapaItens[ei][0];
      var label    = etapaItens[ei][1];
      var offset   = etapaItens[ei][2];
      if (!et[chave] || et[chave] === 'NÃO') {
        var dEtapa  = _addDias(dataAdm, offset);
        var ini9h   = new Date(dEtapa.getFullYear(), dEtapa.getMonth(), dEtapa.getDate(), 9, 0, 0);
        var fim930  = new Date(dEtapa.getFullYear(), dEtapa.getMonth(), dEtapa.getDate(), 9, 30, 0);
        _criarEventoCalendar(
          '⏳ ' + dados.nome + ' — ' + label,
          ini9h, fim930,
          'Pendência admissão | Empresa: ' + dados.empresa + ' | Cargo: ' + (dados.cargo || '')
        );
      }
    }

    return { ok: true, link: dados.tipo === 'digital' ? linkPasta : linkDoc };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  processAtestado
// ───────────────────────────────────────────────────────────
function processAtestado(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Atestados', [
      'Data/Hora', 'Nome Colaborador', 'Empresa', 'Dias', 'CID',
      'Enc. Contab', 'Enc. SST', 'Lanc. Ponto', 'Avis. Lider', 'Link Arquivo'
    ]);

    var agora    = new Date();
    var dataStr  = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var dataNome = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'ddMMyyyy');

    var pastaEmpresa  = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa);
    var pastaColab    = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome);
    var pastaAtestado = _getOrCreateFolder(pastaColab.getId(), 'Atestados');
    var nomeArq       = 'Atestado_' + dados.nome + '_' + dataNome + '.pdf';
    var arquivo       = _salvarArquivo(pastaAtestado, nomeArq, dados.arquivoBase64, dados.arquivoMime);

    sheet.appendRow([
      dataStr, dados.nome, dados.empresa, dados.dias, dados.cid || '',
      dados.encContab || '', dados.encSST || '', dados.encPonto || '', dados.encLider || '',
      arquivo.getUrl()
    ]);

    var diasNum  = parseInt(dados.dias) || 1;
    var atInicio = new Date();
    var atFim    = new Date(atInicio.getTime() + diasNum * 86400000);
    _criarEventoCalendar(
      '🏥 Atestado: ' + dados.nome + ' — ' + dados.empresa + ' (' + dados.dias + 'd)',
      atInicio, atFim,
      'CID: ' + (dados.cid || 'Não informado') + '\nArquivo: ' + arquivo.getUrl()
    );

    return { ok: true, link: arquivo.getUrl() };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  processDesligamento
// ───────────────────────────────────────────────────────────
function processDesligamento(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Desligamentos', [
      'Data/Hora', 'Nome', 'Empresa', 'Data Deslig.', 'Tipo Saida', 'Observacoes', 'Link Arquivo'
    ]);

    var agora    = new Date();
    var dataStr  = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var linkArq  = '';

    if (dados.arquivoBase64) {
      var dataNome      = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'ddMMyyyy');
      var pastaEmpresa  = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa);
      var pastaFunc     = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome);
      var pastaRescisao = _getOrCreateFolder(pastaFunc.getId(), 'Rescisão');
      var nomeArq       = 'Desligamento_' + dados.nome + '_' + dataNome + '.pdf';
      var arquivo       = _salvarArquivo(pastaRescisao, nomeArq, dados.arquivoBase64, dados.arquivoMime);
      linkArq           = arquivo.getUrl();
    }

    sheet.appendRow([
      dataStr, dados.nome, dados.empresa, dados.data || '', dados.tipo || '', dados.obs || '', linkArq
    ]);

    var dataDeslig    = dados.data ? new Date(dados.data + 'T12:00:00') : new Date();
    var dataDesligFim = new Date(dataDeslig.getTime() + 3600000);
    _criarEventoCalendar(
      '⚠️ Rescisão: ' + dados.nome + ' — ' + dados.empresa,
      dataDeslig, dataDesligFim,
      'Tipo: ' + (dados.tipo || '') + '\n' + (dados.obs || '')
    );

    return { ok: true, link: linkArq };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  processCadastroCliente
// ───────────────────────────────────────────────────────────
function processCadastroCliente(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Clientes', [
      'Data/Hora', 'Razao Social', 'CNPJ', 'Contabilidade', 'Contato Contab',
      'Horario', 'Sistema Ponto', 'VR', 'VT', 'Saude', 'Odonto', 'SST', 'Nome Fantasia',
      'Relogio Modelo', 'Relogio Serie', 'WhatsApp Lideranca', 'WhatsApp Equipe'
    ]);

    var dataStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    sheet.appendRow([
      dataStr, dados.razaoSocial, dados.cnpj,
      dados.contabilidade, dados.contabContato || '',
      dados.horario, dados.sistemaPonto || '',
      dados.vr || 'Nao', dados.vt || 'Nao',
      dados.saude || 'Nao', dados.odonto || 'Nao', dados.sst || 'Nao',
      dados.nomeFantasia || dados.razaoSocial,
      dados.relogioModelo || '', dados.relogioSerie || '',
      dados.whatsLideranca || '', dados.whatsEquipe || ''
    ]);

    _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.nomeFantasia || dados.razaoSocial);
    return { ok: true };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  processVaga
// ───────────────────────────────────────────────────────────
function processVaga(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Vagas', [
      'Data/Hora', 'Titulo', 'Empresa', 'Area', 'Nivel', 'Tipo Contrato',
      'Qtd Vagas', 'Motivo', 'Modalidade', 'Cidade', 'Estado',
      'Salario De', 'Salario Ate', 'A Combinar', 'Beneficios',
      'Escolaridade', 'Experiencia', 'Descricao', 'Requisitos', 'Diferenciais',
      'Gestor', 'Data Limite', 'InfoJobs', 'Aprovada', 'Status'
    ]);

    var dataStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    sheet.appendRow([
      dataStr,
      dados.titulo, dados.empresa, dados.area, dados.nivel, dados.tipoContrato,
      dados.qtdVagas || '1', dados.motivo, dados.modalidade,
      dados.cidade || '', dados.estado || '',
      dados.aCombinar ? 'A combinar' : (dados.salarioDe || ''),
      dados.aCombinar ? '' : (dados.salarioAte || ''),
      dados.aCombinar ? 'SIM' : 'NAO',
      dados.beneficios || '',
      dados.escolaridade || '', dados.experiencia || '',
      dados.descricao || '', dados.requisitos || '', dados.diferenciais || '',
      dados.gestor || '', dados.dataLimite || '',
      dados.infojobs || 'Sim',
      dados.aprovada ? 'SIM' : 'PENDENTE',
      'Aberta'
    ]);

    if (dados.dataLimite) {
      var dataLim    = new Date(dados.dataLimite + 'T12:00:00');
      var dataLimFim = new Date(dataLim.getTime() + 3600000);
      _criarEventoCalendar(
        '🔍 Prazo Vaga: ' + dados.titulo + ' — ' + dados.empresa,
        dataLim, dataLimFim,
        'Nível: ' + (dados.nivel || '') + ' | ' + (dados.tipoContrato || '') + '\nGestor: ' + (dados.gestor || '')
      );
    }

    return { ok: true };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  processDocumentosColaborador
// ───────────────────────────────────────────────────────────
function processDocumentosColaborador(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Envio_Docs', [
      'Data/Hora', 'Nome', 'CPF', 'Nascimento', 'Telefone', 'Empresa',
      'LGPD', 'Docs Enviados', 'Link Pasta'
    ]);

    // Cria estrutura de pastas: Drive > Empresa > Colaborador > Admissão
    var pastaEmpresa = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa || 'Sem_Empresa');
    var pastaColab   = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome || 'Colaborador');
    var pastaAdmExt  = _getOrCreateFolder(pastaColab.getId(), 'Admissão');

    var nome    = (dados.nome || 'doc').replace(/\s+/g, '_');
    var docMap  = [
      ['rgFrente',    'RG_Frente'],
      ['rgVerso',     'RG_Verso'],
      ['cpf',         'CPF'],
      ['foto',        'Foto3x4'],
      ['residencia',  'ComprovanteResidencia'],
      ['escolaridade','ComprovanteEscolaridade'],
      ['pis',         'PIS_PASEP'],
      ['titulo',      'TituloEleitor'],
      ['banco',       'DadosBancarios'],
      ['filhos',      'CertidaoFilhos']
    ];

    var enviados = [];
    for (var i = 0; i < docMap.length; i++) {
      var key    = docMap[i][0];
      var label  = docMap[i][1];
      var b64    = dados[key + 'Base64'];
      var mime   = dados[key + 'Mime'];
      if (b64 && mime) {
        _salvarArquivo(pastaAdmExt, label + '_' + nome, b64, mime);
        enviados.push(label);
      }
    }

    var dataStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    sheet.appendRow([
      dataStr,
      dados.nome || '', dados.cpf || '', dados.nascimento || '', dados.telefone || '',
      dados.empresa || '', dados.lgpd ? 'SIM' : 'NÃO',
      enviados.join(', '), pastaAdmExt.getUrl()
    ]);

    var hojeDoc   = new Date();
    var amanhaDoc = new Date(hojeDoc.getTime() + 86400000);
    _criarEventoCalendar(
      '📥 Docs recebidos: ' + dados.nome + ' — ' + (dados.empresa || ''),
      hojeDoc, amanhaDoc,
      'CPF: ' + (dados.cpf || '') + '\nDocs enviados: ' + enviados.join(', ')
    );

    return { ok: true, link: pastaAdmExt.getUrl() };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  processDiagnostico
// ───────────────────────────────────────────────────────────
function processDiagnostico(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Diagnosticos', [
      'Data/Hora', 'Empresa', 'CNPJ', 'Colaboradores', 'Ramo', 'Responsavel', 'Data Diag',
      'Score (%)', 'Nivel Compliance',
      'Q01 Organograma', 'Q02 Desc Cargos', 'Q03 CTPS Regularizada',
      'Q04 Regimento Interno', 'Q05 Sensibilizacao', 'Q06 Canal Denuncias',
      'Q07 PCMSO', 'Q08 PGR', 'Q09 Controle EPI', 'Q10 ASOs em Dia',
      'Q11 Fichas Funcionais', 'Q12 Sistema Ponto', 'Q13 BH Digital',
      'Q14 Intervalos', 'Q15 Controle Ferias', 'Q16 Atestados Arquivados',
      'Q17 Contracheques', 'Q18 Contratos Escritos', 'Q19 LGPD', 'Q20 Acordo BH',
      'Q21 Reconhecimento', 'Q22 Feedback', 'Q23 Reunioes',
      'Observacoes', 'URL Documento'
    ]);

    var r = dados.respostas || {};
    var agora = new Date();
    var dataStr = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var dataFmt = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // ── Gerar documento Google Docs ──
    var docUrl = '';
    try {
      var PERGUNTAS = [
        ['Estrutura e Documentação', [
          ['Q01', 'A empresa possui organograma formalizado?'],
          ['Q02', 'Existem descrições de cargo por função?'],
          ['Q03', 'Todas as CTPs estão regularizadas (registro na CTPS ou e-Social)?'],
          ['Q04', 'Existe Regimento Interno ou Código de Conduta formalizado?'],
          ['Q05', 'Os colaboradores foram sensibilizados sobre normas internas?'],
          ['Q06', 'Existe canal de denúncias ou ouvidoria?']
        ]],
        ['Saúde e Segurança (SST)', [
          ['Q07', 'A empresa possui PCMSO (Programa de Controle Médico de Saúde Ocupacional)?'],
          ['Q08', 'A empresa possui PGR (Programa de Gerenciamento de Riscos)?'],
          ['Q09', 'Existe controle de distribuição e uso de EPIs?'],
          ['Q10', 'Os ASOs (Atestados de Saúde Ocupacional) estão em dia para todos os colaboradores?']
        ]],
        ['Controles Operacionais', [
          ['Q11', 'Os colaboradores possuem fichas funcionais individuais?'],
          ['Q12', 'A empresa utiliza sistema de controle de ponto?'],
          ['Q13', 'O banco de horas é controlado digitalmente?'],
          ['Q14', 'Os intervalos intrajornada são registrados?'],
          ['Q15', 'Existe controle ativo de férias com escala programada?'],
          ['Q16', 'Atestados médicos são arquivados organizadamente?']
        ]],
        ['Contratos e LGPD', [
          ['Q17', 'Os contracheques são assinados pelos colaboradores?'],
          ['Q18', 'Os contratos de trabalho estão formalizados por escrito?'],
          ['Q19', 'A empresa possui política de privacidade / adequação à LGPD?'],
          ['Q20', 'Existe acordo de banco de horas formalmente assinado?']
        ]],
        ['Engajamento e Cultura', [
          ['Q21', 'Existe algum programa de reconhecimento ou incentivo aos colaboradores?'],
          ['Q22', 'Há rotina de feedback entre gestores e colaboradores?'],
          ['Q23', 'A empresa realiza reuniões periódicas de equipe?']
        ]]
      ];

      var nivelColor = { 'Excelente': '#1e7e34', 'Bom': '#155724', 'Médio': '#856404', 'Crítico': '#721c24' };
      var cor = nivelColor[dados.nivel] || '#3D3A34';

      var doc  = DocumentApp.create('Avaliacao de Compliance RH - ' + (dados.empresa || 'Empresa'));
      var body = doc.getBody();
      var tz   = Session.getScriptTimeZone();
      var estBase = {};
      estBase[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
      estBase[DocumentApp.Attribute.FONT_SIZE]   = 11;

      // Cabeçalho
      body.setPageWidth(595.28).setPageHeight(841.89)
          .setMarginTop(50).setMarginBottom(50).setMarginLeft(56).setMarginRight(56);

      var h1 = body.appendParagraph('CRESCER CONSULTORIA EMPRESARIAL');
      h1.setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      h1.editAsText().setFontSize(14).setBold(true).setForegroundColor('#3D3A34');

      var h2 = body.appendParagraph('Avaliação de Compliance em Recursos Humanos');
      h2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      h2.editAsText().setFontSize(12).setBold(false).setItalic(true).setForegroundColor('#5C5952');

      body.appendParagraph(' ');

      // Dados da empresa
      function addLinha(label, valor) {
        var p = body.appendParagraph('');
        var t = p.editAsText();
        t.appendText(label + ': ').setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor('#1C1B18');
        t.appendText(valor || '—').setFontFamily('Arial').setFontSize(10).setBold(false).setForegroundColor('#1C1B18');
      }

      addLinha('Empresa',           dados.empresa || '');
      addLinha('CNPJ',              dados.cnpj || 'Não informado');
      addLinha('Nº de Colaboradores', dados.colaboradores || 'Não informado');
      addLinha('Ramo de Atividade', dados.ramo || 'Não informado');
      addLinha('Responsável',       dados.responsavel || 'Não informado');
      addLinha('Data do Diagnóstico', dados.data || dataFmt);
      addLinha('Realizado por',     'Crescer Consultoria Empresarial');

      body.appendParagraph(' ');

      // Score
      var scoreP = body.appendParagraph('RESULTADO GERAL: ' + dados.scoreTotal + '%  —  ' + (dados.nivel || ''));
      scoreP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      scoreP.editAsText().setFontSize(14).setBold(true).setForegroundColor(cor);

      var descNivel = {
        'Excelente': 'Empresa com excelente nível de conformidade em RH. Manter e evoluir.',
        'Bom':       'Boa conformidade geral. Pontos de atenção devem ser endereçados.',
        'Médio':     'Conformidade parcial. Recomenda-se plano de ação prioritário.',
        'Crítico':   'Conformidade crítica. Ação imediata necessária para mitigar riscos.'
      };
      var descP = body.appendParagraph(descNivel[dados.nivel] || '');
      descP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      descP.editAsText().setFontSize(10).setItalic(true).setForegroundColor(cor);

      body.appendParagraph(' ');

      // Separador
      body.appendHorizontalRule();
      body.appendParagraph(' ');

      // Seções e perguntas
      var qIdx = 1;
      for (var s = 0; s < PERGUNTAS.length; s++) {
        var secNome = PERGUNTAS[s][0];
        var pergs   = PERGUNTAS[s][1];

        var secP = body.appendParagraph(String(s + 1) + '. ' + secNome.toUpperCase());
        secP.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        secP.editAsText().setFontSize(11).setBold(true).setForegroundColor('#3D3A34');

        for (var q = 0; q < pergs.length; q++) {
          var code = pergs[q][0].toLowerCase(); // q01 → q1 fix
          var num  = parseInt(code.replace('q',''), 10);
          var resp = r['q' + num] || 'Não respondido';
          var respColor = resp === 'SIM' ? '#16a34a' : resp === 'NÃO' ? '#dc2626' : '#9B9890';
          var line = body.appendParagraph('');
          var lt   = line.editAsText();
          lt.appendText((num < 10 ? '0' + num : num) + '. ').setFontFamily('Arial').setFontSize(10).setBold(false).setForegroundColor('#5C5952');
          lt.appendText(pergs[q][1]).setFontFamily('Arial').setFontSize(10).setBold(false).setForegroundColor('#1C1B18');
          lt.appendText('   [' + resp + ']').setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(respColor);
          qIdx++;
        }
        body.appendParagraph(' ');
      }

      body.appendHorizontalRule();
      body.appendParagraph(' ');

      // Observações
      if (dados.obs) {
        var obsH = body.appendParagraph('Observações');
        obsH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        obsH.editAsText().setFontSize(11).setBold(true).setForegroundColor('#3D3A34');
        var obsP = body.appendParagraph(dados.obs);
        obsP.editAsText().setFontSize(10).setForegroundColor('#1C1B18');
        body.appendParagraph(' ');
        body.appendHorizontalRule();
        body.appendParagraph(' ');
      }

      // Assinaturas
      var assH = body.appendParagraph('Assinaturas');
      assH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      assH.editAsText().setFontSize(11).setBold(true).setForegroundColor('#3D3A34');

      body.appendParagraph('\n\n__________________________________        __________________________________');
      var ass1 = body.appendParagraph('Consultor(a) Responsável — Crescer        Responsável pela Empresa');
      ass1.editAsText().setFontSize(9).setForegroundColor('#5C5952');
      body.appendParagraph('\nData: _____ / _____ / _______');

      doc.saveAndClose();

      // Mover para pasta da empresa / Diagnóstico
      var pastaEmp  = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa || 'Empresa');
      var pastaDiag = _getOrCreateFolder(pastaEmp.getId(), 'Diagnóstico');
      var docFile   = DriveApp.getFileById(doc.getId());
      docFile.moveTo(pastaDiag);
      docUrl = docFile.getUrl();
    } catch (docErr) {
      // Documento falhou mas salva planilha mesmo assim
      docUrl = '';
    }

    sheet.appendRow([
      dataStr,
      dados.empresa || '', dados.cnpj || '', dados.colaboradores || '',
      dados.ramo || '', dados.responsavel || '', dados.data || '',
      dados.scoreTotal + '%', dados.nivel || '',
      r.q1||'', r.q2||'', r.q3||'', r.q4||'', r.q5||'', r.q6||'',
      r.q7||'', r.q8||'', r.q9||'', r.q10||'',
      r.q11||'', r.q12||'', r.q13||'', r.q14||'', r.q15||'', r.q16||'',
      r.q17||'', r.q18||'', r.q19||'', r.q20||'',
      r.q21||'', r.q22||'', r.q23||'',
      dados.obs || '', docUrl
    ]);

    return { ok: true, score: dados.scoreTotal, nivel: dados.nivel, url: docUrl };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════

// ───────────────────────────────────────────────────────────
//  getColaboradores
// ───────────────────────────────────────────────────────────
function getColaboradores(empresa) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    // Lê de Base_Colaboradores (fonte principal); fallback para Base_Admissoes
    var sheet = ss.getSheetByName('Base_Colaboradores') || ss.getSheetByName('Base_Admissoes');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var h = data[0];
    function ci(name) {
      for (var i = 0; i < h.length; i++) {
        if (h[i].toString().toLowerCase().replace(/[\s\/ã]/g,'') === name.toLowerCase().replace(/[\s\/ã]/g,'')) return i;
      }
      return -1;
    }
    var iNome = ci('nome'), iCpf = ci('cpf'), iCtps = ci('ctps');
    var iNasc = ci('nascimento'), iTel = ci('telefone');
    var iEmp  = ci('empresa'),   iCargo = ci('cargo'), iAdm = ci('dataadmissao');
    var emp   = (empresa || '').toLowerCase().trim();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (emp && (iEmp < 0 || row[iEmp].toString().toLowerCase().trim() !== emp)) continue;
      var nome = iNome >= 0 ? row[iNome].toString().trim() : '';
      if (!nome) continue;
      var dataAdmVal = iAdm >= 0 ? row[iAdm] : '';
      if (dataAdmVal instanceof Date) {
        dataAdmVal = Utilities.formatDate(dataAdmVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      result.push({
        nome:     nome,
        cpf:      iCpf   >= 0 ? row[iCpf].toString().trim()   : '',
        ctps:     iCtps  >= 0 ? row[iCtps].toString().trim()  : '',
        nascimento: iNasc >= 0 ? row[iNasc].toString().trim() : '',
        telefone: iTel   >= 0 ? row[iTel].toString().trim()   : '',
        empresa:  iEmp   >= 0 ? row[iEmp].toString().trim()   : '',
        cargo:    iCargo >= 0 ? row[iCargo].toString().trim()  : '',
        dataAdm:  dataAdmVal  ? dataAdmVal.toString()          : ''
      });
    }
    return result;
  } catch (err) {
    return [];
  }
}

// ───────────────────────────────────────────────────────────
//  gerarDocumento
// ───────────────────────────────────────────────────────────
function gerarDocumento(dados) {
  try {
    var tipo = dados.tipo || '';
    var pastaRaiz     = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_ID);
    var pastaEmpresa  = _getOrCreateFolder(pastaRaiz.getId(), dados.empresa || 'Sem_Empresa');
    var pastaColab    = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome  || 'Colaborador');
    var pastaDocs     = _getOrCreateFolder(pastaColab.getId(), 'Documentos');

    var hoje = new Date();
    var dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var dataArq = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd-MM-yyyy');

    var admStr = '';
    if (dados.dataAdm) {
      try {
        var partes = dados.dataAdm.toString().split('-');
        if (partes.length === 3) admStr = partes[2] + '/' + partes[1] + '/' + partes[0];
        else admStr = dados.dataAdm;
      } catch(e) { admStr = dados.dataAdm; }
    }

    var empDados = _buscarEmpresaDados(dados.empresa);
    var cnpjEmpresa = empDados.cnpj;
    var razaoSocialEmpresa = empDados.razaoSocial;

    var doc, nomeDoc;

    if (tipo === 'lgpd-imagem') {
      nomeDoc = 'TermoLGPD_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
      doc = DocumentApp.create(nomeDoc);
      var body = doc.getBody();
      body.clear();
      var t = body.appendParagraph('TERMO DE CONSENTIMENTO PARA O TRATAMENTO DE DADOS PESSOAIS E USO DE IMAGEM');
      t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t.editAsText().setBold(true).setFontSize(13);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('Eu, ' + dados.nome + ', inscrito(a) no CPF sob o nº ' + (dados.cpf || '_______________') + ', no exercício da função de ' + (dados.cargo || '_______________') + ' na empresa ' + razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : '') + (admStr ? ', admitido(a) em ' + admStr : '') + ', neste ato DECLARO e AUTORIZO:').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      _gerDocSec(body, '1. Do Tratamento de Dados Pessoais (Lei 13.709/2018 — LGPD)');
      body.appendParagraph('Consinto, de forma livre, informada e inequívoca, que a empresa e a CRESCER Consultoria Empresarial coletem, armazenem, utilizem e tratem meus dados pessoais estritamente para fins trabalhistas e previdenciários, incluindo: elaboração de folha de pagamento, registro em eSocial, controle de ponto, cumprimento de obrigações legais e gestão do contrato de trabalho.').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('Declaro estar ciente de que meus dados não serão compartilhados com terceiros, exceto por obrigação legal ou para cumprimento das finalidades acima descritas, e que posso revogar este consentimento a qualquer tempo, mediante solicitação formal, observados os limites legais.').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      _gerDocSec(body, '2. Do Uso de Imagem');
      body.appendParagraph('Autorizo o uso da minha imagem, voz e dados biométricos pela empresa e pela CRESCER Consultoria Empresarial para fins exclusivamente relacionados à relação de trabalho, incluindo: crachá de identificação, registro fotográfico de admissão, sistema de controle de ponto por biometria e comunicações internas da empresa.').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('A presente autorização não abrange uso comercial, publicitário ou divulgação em redes sociais sem consentimento específico e prévio do(a) colaborador(a).').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      _gerDocSec(body, '3. Da Vigência');
      body.appendParagraph('O presente Termo tem validade pelo período em que durar o contrato de trabalho, podendo ser revogado mediante comunicação formal ao RH da CRESCER Consultoria Empresarial.').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
      body.appendParagraph('Assinatura do(a) Colaborador(a)').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

    } else if (tipo === 'banco-horas') {
      nomeDoc = 'BancoHoras_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
      doc = DocumentApp.create(nomeDoc);
      var body = doc.getBody();
      body.clear();
      var t = body.appendParagraph('TERMO DE CIÊNCIA SOBRE A INSTITUIÇÃO DO BANCO DE HORAS');
      t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t.editAsText().setBold(true).setFontSize(13);
      body.appendParagraph('');
      var intro = body.appendParagraph('Eu, ' + dados.nome + ', inscrito no CPF sob o nº ' + dados.cpf + ', no exercício da função de ' + (dados.cargo || '_______________') + ', DECLARO, neste ato, tomar ciência formal sobre a implantação do sistema de Banco de Horas no âmbito da empresa, ' + razaoSocialEmpresa + (cnpjEmpresa ? ', inscrita no CNPJ nº ' + cnpjEmpresa : '') + ', conforme as disposições legais e convencionais aplicáveis.');
      intro.editAsText().setFontSize(11);
      body.appendParagraph('');
      _gerDocSec(body, '1. Da Instituição e Fundamentação');
      body.appendParagraph('O Banco de Horas é instituído com amparo legal no art. 59, §2º da CLT e respaldo coletivo conferido pela Convenção Coletiva de Trabalho 2024/2025, decorrendo a autorização da negociação coletiva firmada entre os sindicatos representativos das categorias econômica e profissional.').editAsText().setFontSize(11);
      body.appendParagraph('');
      _gerDocSec(body, '2. Da Jornada de Trabalho');
      body.appendParagraph('A jornada normal de trabalho é de 44 (quarenta e quatro) horas semanais e 08 horas diárias, podendo ser distribuída conforme a necessidade operacional da empresa.').editAsText().setFontSize(11);
      body.appendParagraph('');
      _gerDocSec(body, '3. Da Vigência');
      body.appendParagraph('O presente regime de Banco de Horas tem vigência de 12 (doze) meses, contados a partir da data de sua implantação, podendo ser renovado mediante nova comunicação formal aos empregados.').editAsText().setFontSize(11);
      body.appendParagraph('');
      body.appendParagraph('');
      var assLinha = body.appendParagraph('_______________________________________________');
      assLinha.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      var assNome = body.appendParagraph(dados.nome);
      assNome.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      assNome.editAsText().setBold(true);
      body.appendParagraph('Funcionário').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph('');
      body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    } else if (tipo === 'aso') {
      var tipoExame = dados.asoTipo || 'Admissional';
      nomeDoc = 'Encaminhamento_ASO_' + tipoExame + '_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
      doc = DocumentApp.create(nomeDoc);
      var body = doc.getBody();
      body.clear();
      var t = body.appendParagraph('ENCAMINHAMENTO PARA EXAME OCUPACIONAL (ASO)');
      t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t.editAsText().setBold(true).setFontSize(13);
      body.appendParagraph('');
      _gerDocCampo(body, 'Data: ', dataStr);
      _gerDocCampo(body, 'Tipo de Exame: ', tipoExame);
      _gerDocCampo(body, 'Empresa: ', razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : ''));
      _gerDocCampo(body, 'Colaborador: ', dados.nome);
      _gerDocCampo(body, 'CPF: ', dados.cpf || '_______________');
      _gerDocCampo(body, 'Cargo: ', dados.cargo || '_______________');
      _gerDocCampo(body, 'Data de Admissão: ', admStr || '_______________');
      body.appendParagraph('');
      body.appendParagraph('Solicitamos a realização do exame médico ocupacional ' + tipoExame.toLowerCase() + ' do(a) colaborador(a) acima identificado(a), conforme determinação da NR-7 — Programa de Controle Médico de Saúde Ocupacional (PCMSO).').editAsText().setFontSize(11);
      body.appendParagraph('');
      body.appendParagraph('');
      body.appendParagraph('___________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph('CRESCER Consultoria Empresarial').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);
      body.appendParagraph('(83) 99670.3518').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    } else if (tipo === 'regulamento') {
      nomeDoc = 'Recibo_Regulamento_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
      doc = DocumentApp.create(nomeDoc);
      var body = doc.getBody();
      body.clear();
      var t = body.appendParagraph('RECIBO DE ENTREGA DO REGULAMENTO INTERNO');
      t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t.editAsText().setBold(true).setFontSize(13);
      body.appendParagraph('');
      body.appendParagraph('Declaro ter recebido um exemplar do Regulamento Interno da empresa ' + razaoSocialEmpresa + ' e me comprometo a cumprir integralmente as normas nele contidas, estando ciente de que o seu desconhecimento não poderá ser alegado como justificativa para qualquer descumprimento.').editAsText().setFontSize(11);
      body.appendParagraph('');
      _gerDocCampo(body, 'Empregado: ', dados.nome);
      _gerDocCampo(body, 'CPF: ', dados.cpf || '_______________');
      _gerDocCampo(body, 'Carteira de Trabalho: ', dados.ctps || '_______________');
      _gerDocCampo(body, 'Cargo: ', dados.cargo || '_______________');
      _gerDocCampo(body, 'Admissão em: ', admStr || '_______________');
      body.appendParagraph('');
      body.appendParagraph('');
      body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph('');
      body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);
      body.appendParagraph('Assinatura do Empregado').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    } else if (tipo === 'declaracao-transporte') {
      var horario = dados.transpHorario || '00:00';
      var valor   = dados.transpValor   || '6,00';
      nomeDoc = 'Declaracao_Transporte_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
      doc = DocumentApp.create(nomeDoc);
      var body = doc.getBody();
      body.clear();
      var t = body.appendParagraph('TERMO DE DECLARAÇÃO E CIÊNCIA');
      t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t.editAsText().setBold(true).setFontSize(13);
      body.appendParagraph('');
      var texto = 'Eu, ' + dados.nome + ', ' + (dados.cargo || 'colaborador(a)') + ', inscrito junto ao Ministério da Fazenda – CPF sob o n. ' + (dados.cpf || '_______________') + (dados.ctps ? ', portador da CTPS n. ' + dados.ctps : '') + (admStr ? ', admitido(a) em ' + admStr : '') + ', declaro ter sido informado pelo meu Empregador – ' + razaoSocialEmpresa.toUpperCase() + ' que:';
      body.appendParagraph(texto).editAsText().setFontSize(11);
      body.appendParagraph('');
      var li1 = body.appendListItem('Em razão da ausência de transporte coletivo urbano no horário de término da jornada, o Empregador pagará ajuda de custo, em dinheiro, para utilização exclusiva com despesas de transporte no deslocamento trabalho-casa;');
      li1.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      li1.editAsText().setFontSize(11);
      var li2 = body.appendListItem('O valor pago pelo Empregador será de R$ ' + valor + ' quando a jornada terminar às ' + horario + '.');
      li2.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      li2.editAsText().setFontSize(11);
      var li3 = body.appendListItem('Os valores serão pagos apenas em dias trabalhados.');
      li3.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      li3.editAsText().setFontSize(11);
      body.appendParagraph('');
      body.appendParagraph('Por fim, declaro estar ciente de que a ajuda de custo poderá ter seu pagamento interrompido caso as linhas de transporte público coletivo urbano passem a estar disponíveis ao final da minha jornada de trabalho.').editAsText().setFontSize(11);
      body.appendParagraph('');
      body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph('');
      body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);

    } else if (tipo === 'recebimento-etica') {
      nomeDoc = 'TermoRecebimento_EticaRegimento_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
      doc = DocumentApp.create(nomeDoc);
      var body = doc.getBody();
      body.clear();
      var t = body.appendParagraph('TERMO DE RECEBIMENTO E COMPROMISSO');
      t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t.editAsText().setBold(true).setFontSize(13);
      var t2 = body.appendParagraph('Código de Ética e Regimento Interno');
      t2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      t2.editAsText().setFontSize(11).setItalic(true);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('Eu, ' + dados.nome + ', portador(a) do CPF nº ' + (dados.cpf || '_______________') + (dados.ctps ? ', CTPS nº ' + dados.ctps : '') + ', admitido(a) em ' + (admStr || '_______________') + ' na empresa ' + dados.empresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : '') + ', no cargo de ' + (dados.cargo || '_______________') + ', DECLARO que:').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      var li1 = body.appendListItem('Recebi, nesta data, um exemplar do Código de Ética e do Regimento Interno da empresa.');
      li1.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li1.editAsText().setFontSize(11);
      var li2 = body.appendListItem('Comprometo-me a ler integralmente o conteúdo dos documentos recebidos.');
      li2.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li2.editAsText().setFontSize(11);
      var li3 = body.appendListItem('Comprometo-me a cumprir todas as normas e diretrizes neles estabelecidas.');
      li3.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li3.editAsText().setFontSize(11);
      var li4 = body.appendListItem('Estou ciente de que o descumprimento das regras poderá acarretar advertências, suspensões ou rescisão por justa causa, conforme a gravidade da infração e o disposto na CLT.');
      li4.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li4.editAsText().setFontSize(11);
      var li5 = body.appendListItem('Declaro que o desconhecimento do conteúdo não poderá ser alegado como justificativa para qualquer descumprimento.');
      li5.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li5.editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
      body.appendParagraph('').editAsText().setFontSize(11);
      body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
      body.appendParagraph('Assinatura do(a) Colaborador(a)').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

    } else {
      return { ok: false, erro: 'Tipo de documento não reconhecido: ' + tipo };
    }

    doc.saveAndClose();
    DriveApp.getFileById(doc.getId()).moveTo(pastaDocs);

    // Registra na planilha
    var ss3 = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var shDoc = _getOrCreateSheet(ss3, 'Documentos_Gerados', ['Data', 'Tipo', 'Empresa', 'Colaborador', 'Link']);
    shDoc.appendRow([dataStr, tipo, dados.empresa, dados.nome, doc.getUrl()]);

    return { ok: true, url: doc.getUrl() };

  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  gerarPacote — gera múltiplos documentos em uma única pasta
// ───────────────────────────────────────────────────────────
function gerarPacote(dados) {
  try {
    var tipos = dados.documentos || [];
    if (!tipos.length) return { ok: false, erro: 'Nenhum documento selecionado.' };

    var pastaRaiz    = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_ID);
    var pastaEmpresa = _getOrCreateFolder(pastaRaiz.getId(), dados.empresa || 'Sem_Empresa');
    var pastaColab   = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome  || 'Colaborador');
    var pastaDocs    = _getOrCreateFolder(pastaColab.getId(), 'Documentos');

    var hoje    = new Date();
    var dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var dataArq = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd-MM-yyyy');

    var admStr = '';
    if (dados.dataAdm) {
      try {
        var partes = dados.dataAdm.toString().split('-');
        if (partes.length === 3) admStr = partes[2] + '/' + partes[1] + '/' + partes[0];
        else admStr = dados.dataAdm;
      } catch(e) { admStr = dados.dataAdm; }
    }

    var empDados2 = _buscarEmpresaDados(dados.empresa);
    var cnpjEmpresa = empDados2.cnpj;
    var razaoSocialEmpresa = empDados2.razaoSocial;

    var gerados = [];
    var ss3 = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var shDoc = _getOrCreateSheet(ss3, 'Documentos_Gerados', ['Data', 'Tipo', 'Empresa', 'Colaborador', 'Link']);

    for (var ti = 0; ti < tipos.length; ti++) {
      var tipo = tipos[ti];
      var doc, nomeDoc, body, t;

      if (tipo === 'lgpd-imagem') {
        nomeDoc = 'TermoLGPD_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
        doc = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('TERMO DE CONSENTIMENTO PARA O TRATAMENTO DE DADOS PESSOAIS E USO DE IMAGEM');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Eu, ' + dados.nome + ', inscrito(a) no CPF sob o nº ' + (dados.cpf || '_______________') + ', no exercício da função de ' + (dados.cargo || '_______________') + ' na empresa ' + razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : '') + (admStr ? ', admitido(a) em ' + admStr : '') + ', neste ato DECLARO e AUTORIZO:').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        _gerDocSec(body, '1. Do Tratamento de Dados Pessoais (Lei 13.709/2018 — LGPD)');
        body.appendParagraph('Consinto, de forma livre, informada e inequívoca, que a empresa e a CRESCER Consultoria Empresarial coletem, armazenem, utilizem e tratem meus dados pessoais estritamente para fins trabalhistas e previdenciários, incluindo: elaboração de folha de pagamento, registro em eSocial, controle de ponto, cumprimento de obrigações legais e gestão do contrato de trabalho.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Declaro estar ciente de que meus dados não serão compartilhados com terceiros, exceto por obrigação legal ou para cumprimento das finalidades acima descritas, e que posso revogar este consentimento a qualquer tempo, mediante solicitação formal, observados os limites legais.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        _gerDocSec(body, '2. Do Uso de Imagem');
        body.appendParagraph('Autorizo o uso da minha imagem, voz e dados biométricos pela empresa e pela CRESCER Consultoria Empresarial para fins exclusivamente relacionados à relação de trabalho, incluindo: crachá de identificação, registro fotográfico de admissão, sistema de controle de ponto por biometria e comunicações internas da empresa.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('A presente autorização não abrange uso comercial, publicitário ou divulgação em redes sociais sem consentimento específico e prévio do(a) colaborador(a).').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        _gerDocSec(body, '3. Da Vigência');
        body.appendParagraph('O presente Termo tem validade pelo período em que durar o contrato de trabalho, podendo ser revogado mediante comunicação formal ao RH da CRESCER Consultoria Empresarial.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('Assinatura do(a) Colaborador(a)').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

      } else if (tipo === 'banco-horas') {
        nomeDoc = 'BancoHoras_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
        doc = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('TERMO DE CIÊNCIA SOBRE A INSTITUIÇÃO DO BANCO DE HORAS');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('');
        body.appendParagraph('Eu, ' + dados.nome + ', inscrito no CPF sob o nº ' + (dados.cpf || '_______________') + ', no exercício da função de ' + (dados.cargo || '_______________') + ', DECLARO, neste ato, tomar ciência formal sobre a implantação do sistema de Banco de Horas no âmbito da empresa, ' + razaoSocialEmpresa + (cnpjEmpresa ? ', inscrita no CNPJ nº ' + cnpjEmpresa : '') + ', conforme as disposições legais e convencionais aplicáveis.').editAsText().setFontSize(11);
        body.appendParagraph('');
        _gerDocSec(body, '1. Da Instituição e Fundamentação');
        body.appendParagraph('O Banco de Horas é instituído com amparo legal no art. 59, §2º da CLT e respaldo coletivo conferido pela Convenção Coletiva de Trabalho 2024/2025, decorrendo a autorização da negociação coletiva firmada entre os sindicatos representativos das categorias econômica e profissional.').editAsText().setFontSize(11);
        body.appendParagraph('');
        _gerDocSec(body, '2. Da Jornada de Trabalho');
        body.appendParagraph('A jornada normal de trabalho é de 44 (quarenta e quatro) horas semanais e 08 horas diárias, podendo ser distribuída conforme a necessidade operacional da empresa.').editAsText().setFontSize(11);
        body.appendParagraph('');
        _gerDocSec(body, '3. Da Vigência');
        body.appendParagraph('O presente regime de Banco de Horas tem vigência de 12 (doze) meses, contados a partir da data de sua implantação, podendo ser renovado mediante nova comunicação formal aos empregados.').editAsText().setFontSize(11);
        body.appendParagraph('');
        body.appendParagraph('');
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);
        body.appendParagraph('Funcionário').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      } else if (tipo === 'aso') {
        var tipoExame = dados.asoTipo || 'Admissional';
        nomeDoc = 'Encaminhamento_ASO_' + tipoExame + '_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
        doc = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('ENCAMINHAMENTO PARA EXAME OCUPACIONAL (ASO)');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('');
        _gerDocCampo(body, 'Data: ', dataStr);
        _gerDocCampo(body, 'Tipo de Exame: ', tipoExame);
        _gerDocCampo(body, 'Empresa: ', razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : ''));
        _gerDocCampo(body, 'Colaborador: ', dados.nome);
        _gerDocCampo(body, 'CPF: ', dados.cpf || '_______________');
        _gerDocCampo(body, 'Cargo: ', dados.cargo || '_______________');
        _gerDocCampo(body, 'Data de Admissão: ', admStr || '_______________');
        body.appendParagraph('');
        body.appendParagraph('Solicitamos a realização do exame médico ocupacional ' + tipoExame.toLowerCase() + ' do(a) colaborador(a) acima identificado(a), conforme determinação da NR-7 — Programa de Controle Médico de Saúde Ocupacional (PCMSO).').editAsText().setFontSize(11);
        body.appendParagraph('');
        body.appendParagraph('');
        body.appendParagraph('___________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('CRESCER Consultoria Empresarial').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);
        body.appendParagraph('(83) 99670.3518').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      } else if (tipo === 'regulamento') {
        nomeDoc = 'Recibo_Regulamento_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
        doc = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('RECIBO DE ENTREGA DO REGULAMENTO INTERNO');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('');
        body.appendParagraph('Declaro ter recebido um exemplar do Regulamento Interno da empresa ' + razaoSocialEmpresa + ' e me comprometo a cumprir integralmente as normas nele contidas, estando ciente de que o seu desconhecimento não poderá ser alegado como justificativa para qualquer descumprimento.').editAsText().setFontSize(11);
        body.appendParagraph('');
        _gerDocCampo(body, 'Empregado: ', dados.nome);
        _gerDocCampo(body, 'CPF: ', dados.cpf || '_______________');
        _gerDocCampo(body, 'Carteira de Trabalho: ', dados.ctps || '_______________');
        _gerDocCampo(body, 'Cargo: ', dados.cargo || '_______________');
        _gerDocCampo(body, 'Admissão em: ', admStr || '_______________');
        body.appendParagraph('');
        body.appendParagraph('');
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);
        body.appendParagraph('Assinatura do Empregado').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      } else if (tipo === 'recebimento-etica') {
        nomeDoc = 'TermoRecebimento_EticaRegimento_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
        doc = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('TERMO DE RECEBIMENTO E COMPROMISSO');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        var t2 = body.appendParagraph('Código de Ética e Regimento Interno');
        t2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t2.editAsText().setFontSize(11).setItalic(true);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Eu, ' + dados.nome + ', portador(a) do CPF nº ' + (dados.cpf || '_______________') + (dados.ctps ? ', CTPS nº ' + dados.ctps : '') + ', admitido(a) em ' + (admStr || '_______________') + ' na empresa ' + razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : '') + ', no cargo de ' + (dados.cargo || '_______________') + ', DECLARO que:').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        var li1 = body.appendListItem('Recebi, nesta data, um exemplar do Código de Ética e do Regimento Interno da empresa.');
        li1.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li1.editAsText().setFontSize(11);
        var li2 = body.appendListItem('Comprometo-me a ler integralmente o conteúdo dos documentos recebidos.');
        li2.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li2.editAsText().setFontSize(11);
        var li3 = body.appendListItem('Comprometo-me a cumprir todas as normas e diretrizes neles estabelecidas.');
        li3.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li3.editAsText().setFontSize(11);
        var li4 = body.appendListItem('Estou ciente de que o descumprimento das regras poderá acarretar advertências, suspensões ou rescisão por justa causa, conforme a gravidade da infração e o disposto na CLT.');
        li4.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li4.editAsText().setFontSize(11);
        var li5 = body.appendListItem('Declaro que o desconhecimento do conteúdo não poderá ser alegado como justificativa para qualquer descumprimento.');
        li5.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li5.editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('Assinatura do(a) Colaborador(a)').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

      } else if (tipo === 'declaracao-transporte') {
        var horario = dados.transpHorario || '00:00';
        var valor   = dados.transpValor   || '6,00';
        nomeDoc = 'Declaracao_Transporte_' + dados.nome.replace(/\s+/g,'_') + '_' + dataArq;
        doc = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('TERMO DE DECLARAÇÃO E CIÊNCIA');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('');
        var texto = 'Eu, ' + dados.nome + ', ' + (dados.cargo || 'colaborador(a)') + ', inscrito junto ao Ministério da Fazenda – CPF sob o n. ' + (dados.cpf || '_______________') + (dados.ctps ? ', portador da CTPS n. ' + dados.ctps : '') + (admStr ? ', admitido(a) em ' + admStr : '') + ', declaro ter sido informado pelo meu Empregador – ' + razaoSocialEmpresa.toUpperCase() + ' que:';
        body.appendParagraph(texto).editAsText().setFontSize(11);
        body.appendParagraph('');
        var li1 = body.appendListItem('Em razão da ausência de transporte coletivo urbano no horário de término da jornada, o Empregador pagará ajuda de custo, em dinheiro, para utilização exclusiva com despesas de transporte no deslocamento trabalho-casa;');
        li1.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li1.editAsText().setFontSize(11);
        var li2 = body.appendListItem('O valor pago pelo Empregador será de R$ ' + valor + ' quando a jornada terminar às ' + horario + '.');
        li2.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li2.editAsText().setFontSize(11);
        var li3 = body.appendListItem('Os valores serão pagos apenas em dias trabalhados.');
        li3.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER); li3.editAsText().setFontSize(11);
        body.appendParagraph('');
        body.appendParagraph('Por fim, declaro estar ciente de que a ajuda de custo poderá ter seu pagamento interrompido caso as linhas de transporte público coletivo urbano passem a estar disponíveis ao final da minha jornada de trabalho.').editAsText().setFontSize(11);
        body.appendParagraph('');
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true);

      } else {
        continue; // tipo desconhecido, pula
      }

      doc.saveAndClose();
      DriveApp.getFileById(doc.getId()).moveTo(pastaDocs);
      shDoc.appendRow([dataStr, tipo, dados.empresa, dados.nome, doc.getUrl()]);
      gerados.push({ tipo: tipo, nome: nomeDoc, url: doc.getUrl() });
    }

    if (!gerados.length) return { ok: false, erro: 'Nenhum documento pôde ser gerado.' };

    return { ok: true, pasta: pastaDocs.getUrl(), documentos: gerados };

  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// Busca CNPJ e Razão Social completa a partir do nome fantasia ou razão social
function _buscarEmpresaDados(empresa) {
  var resultado = { cnpj: '', razaoSocial: empresa || '' };
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sh = ss.getSheetByName('Base_Clientes');
    if (!sh) return resultado;
    var d = sh.getDataRange().getValues();
    var busca = (empresa || '').toLowerCase().trim();
    for (var i = 1; i < d.length; i++) {
      var nomeF = (d[i][12] || d[i][1] || '').toString().toLowerCase().trim();
      var razS  = (d[i][1]  || '').toString().toLowerCase().trim();
      if (nomeF === busca || razS === busca) {
        resultado.cnpj       = d[i][2] ? d[i][2].toString() : '';
        resultado.razaoSocial = d[i][1] ? d[i][1].toString() : (empresa || '');
        break;
      }
    }
  } catch(e) {}
  return resultado;
}

function _gerDocSec(body, titulo) {
  var p = body.appendParagraph(titulo);
  p.editAsText().setBold(true).setFontSize(11);
  return p;
}

function _gerDocCampo(body, label, valor) {
  var p = body.appendParagraph('');
  var txt = p.editAsText();
  var valStr = valor || '';
  txt.insertText(0, label + valStr);
  txt.setBold(0, label.length - 1, true);
  if (valStr.length > 0) {
    txt.setBold(label.length, label.length + valStr.length - 1, false);
  }
  txt.setFontSize(0, label.length + valStr.length - 1, 11);
}

// ───────────────────────────────────────────────────────────
//  criarDocJornada
// ───────────────────────────────────────────────────────────
function criarDocJornada(dados) {
  try {
    // ── Pastas: Drive > Empresa > Colaborador > Jornada ──
    var pastaEmpresa  = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa || 'Sem_Empresa');
    var pastaColabJ   = _getOrCreateFolder(pastaEmpresa.getId(), dados.colaborador || 'Colaborador');
    var pastaJornada  = _getOrCreateFolder(pastaColabJ.getId(), 'Jornada');

    // ── Cria o documento ──
    var nomeDoc = 'Jornada_' + (dados.colaborador || 'Colaborador').replace(/\s+/g, '_') + '_' + (dados.dataGeracao || '').replace(/\//g, '-');
    var doc  = DocumentApp.create(nomeDoc);
    var body = doc.getBody();
    body.clear();

    // ── Estilos base ──
    var estilo = {};
    estilo[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    estilo[DocumentApp.Attribute.FONT_SIZE]   = 11;
    body.setAttributes(estilo);

    // Título
    var titulo = body.appendParagraph('RELATÓRIO DE MONITORAMENTO DE JORNADA E ASSIDUIDADE');
    titulo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titulo.editAsText().setFontSize(0, titulo.getText().length - 1, 13).setBold(0, titulo.getText().length - 1, true);

    body.appendParagraph('');

    // Cabeçalho
    _jornadaParBold(body, 'Data: ',        dados.dataGeracao  || '');
    _jornadaParBold(body, 'Assunto: ',     'Relatório de Monitoramento de Jornada e Assiduidade');
    _jornadaParBold(body, 'Colaborador: ', dados.colaborador  || '');
    var periodo = (dados.periodoIni && dados.periodoFim) ? dados.periodoIni + ' a ' + dados.periodoFim : (dados.periodoIni || '');
    _jornadaParBold(body, 'Período: ',     periodo);

    body.appendHorizontalRule();
    body.appendParagraph('');

    // ── Dados por mês ──
    var meses = dados.meses || [];
    meses.forEach(function(mes) {
      // Título do mês
      var mesPar = body.appendParagraph(mes.key || '');
      mesPar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      mesPar.editAsText().setBold(true).setFontSize(12);

      body.appendParagraph('');

      // Atrasos
      if (mes.atrasos && mes.atrasos.length) {
        var labelA = body.appendParagraph('Atrasos de entrada:');
        labelA.editAsText().setBold(true);

        mes.atrasos.forEach(function(a) {
          var atrasoTexto = _atrasoStrGs(a.atraso);
          var prefixo = a.diaShort + ': Entrada às ' + a.entrada + ' – Atraso de ';
          var li = body.appendListItem(prefixo);
          li.setGlyphType(DocumentApp.GlyphType.BULLET);
          li.appendText(atrasoTexto).setForegroundColor('#c0392b').setBold(true);
        });
      } else {
        var okA = body.appendParagraph('Sem atrasos de chegada neste mês.');
        okA.editAsText().setForegroundColor('#666666').setItalic(true);
      }

      body.appendParagraph('');

      // Intervalos
      if (mes.intervalos && mes.intervalos.length) {
        var labelI = body.appendParagraph('Descumprimentos de intervalo:');
        labelI.editAsText().setBold(true);

        mes.intervalos.forEach(function(iv) {
          var ivStr   = _minToHoraGs(iv.ivMin);
          var prefixo = iv.diaShort + ': Fez ' + ivStr + ' de intervalo (Saiu ' + iv.saida + ', voltou ' + iv.retorno + ') ';
          var li = body.appendListItem(prefixo);
          li.setGlyphType(DocumentApp.GlyphType.BULLET);
          if (iv.tipo === 'cedo') {
            li.appendText('— Descumpriu (Voltou mais cedo)');
          } else {
            li.appendText('– Estouro de ');
            li.appendText(_atrasoStrGs(iv.diff)).setForegroundColor('#c0392b').setBold(true);
          }
        });
      } else {
        var okI = body.appendParagraph('Sem descumprimentos de intervalo neste mês.');
        okI.editAsText().setForegroundColor('#666666').setItalic(true);
      }

      body.appendParagraph('');
    });

    body.appendHorizontalRule();
    body.appendParagraph('');

    // ── Resumo Geral ──
    var t = dados.totais || {};
    var resumoTit = body.appendParagraph('RESUMO GERAL');
    resumoTit.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    resumoTit.editAsText().setBold(true).setFontSize(12);
    body.appendParagraph('');

    if ((t.faltas || 0) > 0) {
      var pF = body.appendListItem('Faltas / ausências sem registro: ');
      pF.setGlyphType(DocumentApp.GlyphType.BULLET);
      pF.appendText(t.faltas + ' dia' + (t.faltas > 1 ? 's' : '')).setForegroundColor('#c0392b').setBold(true);
    }
    if ((t.atrasos || 0) > 0) {
      var pA = body.appendListItem('Atrasos de entrada: ');
      pA.setGlyphType(DocumentApp.GlyphType.BULLET);
      pA.appendText(t.atrasos + ' ocorrência' + (t.atrasos > 1 ? 's' : '') + ', acumulando ' + _atrasoStrGs(t.atrasoMin || 0)).setForegroundColor('#c0392b').setBold(true);
    }
    if ((t.antecipadas || 0) > 0) {
      var pS = body.appendListItem('Saídas antecipadas: ');
      pS.setGlyphType(DocumentApp.GlyphType.BULLET);
      pS.appendText(t.antecipadas + ' ocorrência' + (t.antecipadas > 1 ? 's' : '')).setForegroundColor('#c0392b').setBold(true);
    }
    if ((t.extraMin || 0) > 0) {
      var pEx = body.appendListItem('Horas extras registradas: ');
      pEx.setGlyphType(DocumentApp.GlyphType.BULLET);
      pEx.appendText(_atrasoStrGs(t.extraMin) + ' no total').setForegroundColor('#1a7a4a').setBold(true);
    }
    var cedoCount = (t.iv || 0) - (t.estouros || 0);
    if (cedoCount > 0) {
      var pC = body.appendListItem('Retornos antecipados do intervalo: ');
      pC.setGlyphType(DocumentApp.GlyphType.BULLET);
      pC.appendText(cedoCount + ' ocorrência' + (cedoCount > 1 ? 's' : '')).setForegroundColor('#c0392b').setBold(true);
    }
    if ((t.estouros || 0) > 0) {
      var pE = body.appendListItem('Estouros de intervalo: ');
      pE.setGlyphType(DocumentApp.GlyphType.BULLET);
      pE.appendText(t.estouros + ' ocorrência' + (t.estouros > 1 ? 's' : '')).setForegroundColor('#c0392b').setBold(true);
    }
    if ((t.incompletos || 0) > 0) {
      var pI = body.appendListItem('Dias com registro incompleto (sem saída): ');
      pI.setGlyphType(DocumentApp.GlyphType.BULLET);
      pI.appendText(t.incompletos + ' dia' + (t.incompletos > 1 ? 's' : '')).setForegroundColor('#c0392b').setBold(true);
    }

    body.appendParagraph('');
    body.appendHorizontalRule();
    body.appendParagraph('');

    // ── Análise do Comportamento ──
    var analTit = body.appendParagraph('ANÁLISE DO COMPORTAMENTO DE JORNADA');
    analTit.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    analTit.editAsText().setBold(true).setFontSize(12);
    body.appendParagraph('');

    if (dados.observacoes && dados.observacoes.trim()) {
      var obsP = body.appendParagraph(dados.observacoes.trim());
      obsP.editAsText().setFontSize(11);
    } else {
      var obsP2 = body.appendParagraph('[Adicionar análise da consultora]');
      obsP2.editAsText().setItalic(true).setForegroundColor('#999999');
    }

    body.appendParagraph('');

    // ── Assinatura ──
    for (var s = 0; s < 3; s++) body.appendParagraph('');
    var linha = body.appendParagraph('___________________________');
    linha.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var nome  = body.appendParagraph('Alanne Oliveira');
    nome.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    nome.editAsText().setBold(true).setFontSize(12);
    var emp   = body.appendParagraph('CRESCER Consultoria Empresarial');
    emp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var tel   = body.appendParagraph('(83) 99670.3518');
    tel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();

    // ── Move para pasta de jornada do colaborador ──
    DriveApp.getFileById(doc.getId()).moveTo(pastaJornada);

    // ── Registra na planilha ──
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Jornadas', ['Data', 'Empresa', 'Colaborador', 'Período', 'Link do Documento']);
    sheet.appendRow([
      dados.dataGeracao || '',
      dados.empresa     || '',
      dados.colaborador || '',
      periodo,
      doc.getUrl()
    ]);

    return { ok: true, url: doc.getUrl(), id: doc.getId() };

  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

function _jornadaParBold(body, label, valor) {
  var par = body.appendParagraph('');
  var txt = par.editAsText();
  var valStr = valor || '';
  txt.insertText(0, label + valStr);
  txt.setBold(0, label.length - 1, true);
  if (valStr.length > 0) {
    txt.setBold(label.length, label.length + valStr.length - 1, false);
  }
}

function _minToHoraGs(min) {
  var m = Math.abs(min || 0);
  var h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' : '') + h + 'h' + (mm < 10 ? '0' : '') + mm;
}

function _atrasoStrGs(min) {
  var m = Math.abs(min || 0);
  if (m < 60) return m + ' minuto' + (m !== 1 ? 's' : '');
  var h = Math.floor(m / 60), mm = m % 60;
  var s = (h < 10 ? '0' : '') + h + 'h';
  if (mm > 0) s += (mm < 10 ? '0' : '') + mm;
  return s + ' minutos';
}

function _getOrCreateSheet(ss, nome, cabecalho) {
  var sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cabecalho);
    sheet.getRange(1, 1, 1, cabecalho.length)
      .setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _getOrCreateFolder(parentId, nome) {
  var parent = DriveApp.getFolderById(parentId);
  var iter   = parent.getFoldersByName(nome);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(nome);
}

function _salvarArquivo(pasta, nome, base64, mimeType) {
  var mime  = mimeType || 'application/pdf';
  var raw   = base64.indexOf(',') !== -1 ? base64.split(',')[1] : base64;
  var bytes = Utilities.base64Decode(raw);
  var blob  = Utilities.newBlob(bytes, mime, nome);

  if (mime.indexOf('image/') === 0) {
    var docTemp = DocumentApp.create('_temp_' + nome);
    docTemp.getBody().appendImage(blob);
    docTemp.saveAndClose();
    var docFile = DriveApp.getFileById(docTemp.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(nome.replace(/\.[^.]+$/, '.pdf'));
    docFile.setTrashed(true);
    return pasta.createFile(pdfBlob);
  }
  return pasta.createFile(blob);
}

function _gerarDocFisico(dados, agora) {
  try {
    var templateFile = DriveApp.getFileById(CONFIG.DOCS_TEMPLATE_ID);
    var pastaEmpresa = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa);
    var pastaFunc    = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome);
    var pastaAdmFis  = _getOrCreateFolder(pastaFunc.getId(), 'Admissão');
    var copia        = templateFile.makeCopy('Admissao Fisica - ' + dados.nome, pastaAdmFis);
    var doc          = DocumentApp.openById(copia.getId());
    var body         = doc.getBody();
    body.replaceText('{{NOME}}',    dados.nome);
    body.replaceText('{{CPF}}',     dados.cpf);
    body.replaceText('{{EMPRESA}}', dados.empresa);
    body.replaceText('{{CARGO}}',   dados.cargo || '');
    body.replaceText('{{DATA}}',    Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
    body.replaceText('{{NASC}}',    dados.nascimento || '');
    body.replaceText('{{TEL}}',     dados.telefone || '');
    doc.saveAndClose();
    return copia.getUrl();
  } catch (err) {
    return _gerarDocBasico(dados, agora);
  }
}

function _criarEventoCalendar(titulo, inicio, fim, descricao) {
  try {
    var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!cal) return;
    cal.createEvent(titulo, inicio, fim, { description: descricao || '' });
  } catch (e) {
    // silencioso — nao bloqueia o processo principal
  }
}

// Adiciona N dias corridos a uma data (n pode ser negativo)
function _addDias(data, n) {
  var d = new Date(data.getTime());
  d.setDate(d.getDate() + n);
  return d;
}

function salvarAtividade(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Atividades', ['Data Atividade', 'Titulo', 'Observacao', 'Data Registro']);
    var tz    = Session.getScriptTimeZone();
    var agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    var dataAtivFmt = '';
    try {
      var p = (dados.data || '').split('-');
      if (p.length === 3) dataAtivFmt = p[2] + '/' + p[1] + '/' + p[0] + ' 08:00';
      else dataAtivFmt = dados.data || '';
    } catch(e) { dataAtivFmt = dados.data || ''; }
    sheet.appendRow([dataAtivFmt, dados.titulo || '', dados.obs || '', agora]);
    return { ok: true };
  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  gerarPacoteDes — documentos de rescisão (aviso prévio, pedido demissão, ASO demissional)
// ───────────────────────────────────────────────────────────
function gerarPacoteDes(dados) {
  try {
    var tipos = dados.documentos || [];
    if (!tipos.length) return { ok: false, erro: 'Nenhum documento selecionado.' };

    var empDados           = _buscarEmpresaDados(dados.empresa);
    var cnpjEmpresa        = empDados.cnpj;
    var razaoSocialEmpresa = empDados.razaoSocial;

    // Buscar CPF/cargo do colaborador em Base_Colaboradores
    var cpfColab = '', cargoColab = '';
    try {
      var ss0 = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      var shC = ss0.getSheetByName('Base_Colaboradores');
      if (shC) {
        var dc   = shC.getDataRange().getValues();
        var nbk  = (dados.nome || '').toLowerCase().trim();
        for (var ci = 1; ci < dc.length; ci++) {
          if ((dc[ci][2] || '').toString().toLowerCase().trim() === nbk) {
            cpfColab   = dc[ci][3] ? dc[ci][3].toString() : '';
            cargoColab = dc[ci][7] ? dc[ci][7].toString() : '';
            break;
          }
        }
      }
    } catch(e2) {}

    var pastaRaiz    = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_ID);
    var pastaEmpresa = _getOrCreateFolder(pastaRaiz.getId(), dados.empresa || 'Sem_Empresa');
    var pastaColab   = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome  || 'Colaborador');
    var pastaDocs    = _getOrCreateFolder(pastaColab.getId(), 'Rescisão');

    var tz      = Session.getScriptTimeZone();
    var hoje    = new Date();
    var dataStr = Utilities.formatDate(hoje, tz, 'dd/MM/yyyy');
    var dataArq = Utilities.formatDate(hoje, tz, 'dd-MM-yyyy');

    var dataDesStr = '';
    if (dados.data) {
      var pD = dados.data.toString().split('-');
      if (pD.length === 3) dataDesStr = pD[2] + '/' + pD[1] + '/' + pD[0];
      else dataDesStr = dados.data;
    }

    var gerados = [];
    var ss2   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var shDoc = _getOrCreateSheet(ss2, 'Documentos_Gerados', ['Data', 'Tipo', 'Empresa', 'Colaborador', 'Link']);

    for (var ti = 0; ti < tipos.length; ti++) {
      var tipo = tipos[ti];
      var doc, nomeDoc, body, t;

      if (tipo === 'aviso-previo') {
        nomeDoc = 'AvisoPrevio_' + (dados.nome || '').replace(/\s+/g,'_') + '_' + dataArq;
        doc  = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('AVISO PRÉVIO');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('').editAsText().setFontSize(11);
        _gerDocCampo(body, 'Data: ', dataStr);
        _gerDocCampo(body, 'Empresa: ', razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : ''));
        _gerDocCampo(body, 'Colaborador: ', dados.nome || '');
        _gerDocCampo(body, 'CPF: ', cpfColab || '_______________');
        _gerDocCampo(body, 'Cargo: ', cargoColab || '_______________');
        _gerDocCampo(body, 'Data de Desligamento: ', dataDesStr || '_______________');
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Prezado(a) ' + (dados.nome || '') + ',').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Comunicamos que V.Sa. está sendo dispensado(a) do quadro de colaboradores desta empresa, sem justa causa, nos termos do art. 487 da Consolidação das Leis do Trabalho (CLT).').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('O aviso prévio tem início na data acima indicada. A modalidade de cumprimento (trabalhado ou indenizado) será comunicada pelo RH da CRESCER Consultoria Empresarial.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(razaoSocialEmpresa).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('Representante da Empresa').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome || '').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('Colaborador(a) — Ciente').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

      } else if (tipo === 'pedido-demissao') {
        nomeDoc = 'PedidoDemissao_' + (dados.nome || '').replace(/\s+/g,'_') + '_' + dataArq;
        doc  = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('PEDIDO DE DEMISSÃO');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Campina Grande, ' + dataStr).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Eu, ' + (dados.nome || '') + ', inscrito(a) no CPF sob o nº ' + (cpfColab || '_______________') + ', ocupando o cargo de ' + (cargoColab || '_______________') + ' na empresa ' + razaoSocialEmpresa + (cnpjEmpresa ? ', inscrita no CNPJ nº ' + cnpjEmpresa : '') + ', venho por meio desta solicitar formalmente meu desligamento voluntário desta empresa, com data prevista para ' + (dataDesStr || '_______________') + '.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Declaro estar ciente das implicações legais do pedido de demissão, incluindo as regras relativas ao aviso prévio e às verbas rescisórias aplicáveis.').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(dados.nome || '').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('CPF: ' + (cpfColab || '_______________')).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(razaoSocialEmpresa).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('Testemunha / RH — CRESCER Consultoria Empresarial').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

      } else if (tipo === 'aso-demissional') {
        nomeDoc = 'Encaminhamento_ASO_Demissional_' + (dados.nome || '').replace(/\s+/g,'_') + '_' + dataArq;
        doc  = DocumentApp.create(nomeDoc);
        body = doc.getBody(); body.clear();
        t = body.appendParagraph('ENCAMINHAMENTO PARA EXAME OCUPACIONAL (ASO)');
        t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        t.editAsText().setBold(true).setFontSize(13);
        body.appendParagraph('').editAsText().setFontSize(11);
        _gerDocCampo(body, 'Data: ', dataStr);
        _gerDocCampo(body, 'Tipo de Exame: ', 'Demissional');
        _gerDocCampo(body, 'Empresa: ', razaoSocialEmpresa + (cnpjEmpresa ? ' — CNPJ: ' + cnpjEmpresa : ''));
        _gerDocCampo(body, 'Colaborador: ', dados.nome || '');
        _gerDocCampo(body, 'CPF: ', cpfColab || '_______________');
        _gerDocCampo(body, 'Cargo: ', cargoColab || '_______________');
        _gerDocCampo(body, 'Data de Desligamento: ', dataDesStr || '_______________');
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('Solicitamos a realização do exame médico ocupacional demissional do(a) colaborador(a) acima identificado(a), conforme determinação da NR-7 — Programa de Controle Médico de Saúde Ocupacional (PCMSO).').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('').editAsText().setFontSize(11);
        body.appendParagraph('___________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('CRESCER Consultoria Empresarial').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
        body.appendParagraph('(83) 99670.3518').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);

      } else {
        continue;
      }

      doc.saveAndClose();
      var file = DriveApp.getFileById(doc.getId());
      file.moveTo(pastaDocs);
      var url = file.getUrl();
      gerados.push({ tipo: tipo, nome: nomeDoc, url: url });
      shDoc.appendRow([dataStr, tipo, dados.empresa || '', dados.nome || '', url]);
    }

    return { ok: true, pasta: pastaDocs.getUrl(), documentos: gerados };
  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

// ───────────────────────────────────────────────────────────
//  salvarEntrevistaDesligamento
// ───────────────────────────────────────────────────────────
function salvarEntrevistaDesligamento(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Entrevistas_Desligamento', [
      'Data Registro', 'Empresa', 'Colaborador', 'Cargo', 'Data Entrevista', 'Tipo Saída',
      'Motivo', 'Ambiente (1-5)', 'Liderança (1-5)', 'Crescimento', 'Recomenda',
      'O que mudar', 'Obs RH', 'Link Doc'
    ]);

    var empDados           = _buscarEmpresaDados(dados.empresa);
    var cnpjEmpresa        = empDados.cnpj;
    var razaoSocialEmpresa = empDados.razaoSocial;

    var cpfColab = '', cargoColab = '';
    try {
      var sh  = ss.getSheetByName('Base_Colaboradores');
      if (sh) {
        var d  = sh.getDataRange().getValues();
        var nb = (dados.nome || '').toLowerCase().trim();
        for (var i = 1; i < d.length; i++) {
          if ((d[i][2] || '').toString().toLowerCase().trim() === nb) {
            cpfColab   = d[i][3] ? d[i][3].toString() : '';
            cargoColab = d[i][7] ? d[i][7].toString() : '';
            break;
          }
        }
      }
    } catch(e2) {}

    var tz    = Session.getScriptTimeZone();
    var agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    var dataArq = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy');
    var dataEntStr = '';
    if (dados.data) {
      var pE = dados.data.toString().split('-');
      if (pE.length === 3) dataEntStr = pE[2] + '/' + pE[1] + '/' + pE[0];
      else dataEntStr = dados.data;
    }

    // Gerar documento no Drive
    var pastaRaiz    = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_ID);
    var pastaEmpresa = _getOrCreateFolder(pastaRaiz.getId(), dados.empresa || 'Sem_Empresa');
    var pastaColab   = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome  || 'Colaborador');
    var pastaDocs    = _getOrCreateFolder(pastaColab.getId(), 'Rescisão');
    var nomeDoc = 'EntrevistaDesligamento_' + (dados.nome || '').replace(/\s+/g,'_') + '_' + dataArq;
    var doc  = DocumentApp.create(nomeDoc);
    var body = doc.getBody(); body.clear();
    var t = body.appendParagraph('ENTREVISTA DE DESLIGAMENTO');
    t.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    t.editAsText().setBold(true).setFontSize(13);
    body.appendParagraph('').editAsText().setFontSize(11);
    _gerDocCampo(body, 'Data da Entrevista: ', dataEntStr || agora.split(' ')[0]);
    _gerDocCampo(body, 'Empresa: ', razaoSocialEmpresa);
    _gerDocCampo(body, 'Colaborador: ', dados.nome || '');
    _gerDocCampo(body, 'Cargo: ', cargoColab || '_______________');
    _gerDocCampo(body, 'Tipo de Saída: ', dados.tipo || '');
    body.appendParagraph('').editAsText().setFontSize(11);
    _gerDocSec(body, 'Respostas do Colaborador');
    _gerDocCampo(body, '1. Motivo principal da saída: ', dados.motivo || '—');
    _gerDocCampo(body, '2. Ambiente de trabalho: ', (dados.amb ? dados.amb + '/5' : '—'));
    _gerDocCampo(body, '3. Relação com a liderança: ', (dados.lid ? dados.lid + '/5' : '—'));
    _gerDocCampo(body, '4. Oportunidades de crescimento: ', dados.crescimento || '—');
    _gerDocCampo(body, '5. Recomendaria a empresa: ', dados.recomenda || '—');
    body.appendParagraph('').editAsText().setFontSize(11);
    body.appendParagraph('6. O que poderia ter sido diferente:').editAsText().setBold(true).setFontSize(11);
    body.appendParagraph(dados.mudanca || '—').editAsText().setFontSize(11);
    body.appendParagraph('').editAsText().setFontSize(11);
    _gerDocSec(body, 'Observações do RH');
    body.appendParagraph(dados.obs || '—').editAsText().setFontSize(11);
    body.appendParagraph('').editAsText().setFontSize(11);
    body.appendParagraph('').editAsText().setFontSize(11);
    body.appendParagraph('_______________________________________________').setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('CRESCER Consultoria Empresarial').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
    body.appendParagraph('Consultora Responsável').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(11);
    doc.saveAndClose();
    var file = DriveApp.getFileById(doc.getId());
    file.moveTo(pastaDocs);
    var url = file.getUrl();

    sheet.appendRow([
      agora, dados.empresa || '', dados.nome || '', cargoColab, dataEntStr,
      dados.tipo || '', dados.motivo || '', dados.amb || '', dados.lid || '',
      dados.crescimento || '', dados.recomenda || '', dados.mudanca || '', dados.obs || '', url
    ]);

    return { ok: true, url: url };
  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

function getAtividades() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var fontes = [
      ['Base_Envio_Docs',    'Docs Admissão',  'ok'],
      ['Base_Atestados',     'Atestado',       'warn'],
      ['Base_Desligamentos', 'Desligamento',   'err'],
      ['Base_Vagas',         'Vaga',           'gold'],
      ['Base_Admissoes',     'Admissão',       'ok'],
      ['Base_Atividades',    'Atividade',      'gold']
    ];
    var lista = [];
    for (var f = 0; f < fontes.length; f++) {
      var sheet = ss.getSheetByName(fontes[f][0]);
      if (!sheet) continue;
      var rows  = sheet.getDataRange().getValues();
      var total = rows.length;
      for (var i = Math.max(1, total - 6); i < total; i++) {
        if (!rows[i][0]) continue;
        lista.push({
          tipo:    fontes[f][1],
          cor:     fontes[f][2],
          data:    String(rows[i][0]),
          nome:    String(rows[i][1] || ''),
          empresa: String(rows[i][2] || '')
        });
      }
    }
    // Incluir todos os lembretes pendentes (inclusive futuros) para pontos no calendário
    var todosLems = _todosLembretesPendentes(ss);
    todosLems.forEach(function(l) {
      lista.push({ tipo: 'Lembrete', cor: 'gold', data: l.dataAlvo, nome: l.titulo, empresa: l.empresa || '' });
    });
    // Incluir aniversários próximos
    var aniv = _aniversariantesProximos(ss);
    aniv.forEach(function(a) {
      lista.push({ tipo: 'Aniversário', cor: 'ok', data: a.dataAniv, nome: a.nome, empresa: a.empresa || '' });
    });
    lista.sort(function(a, b) { return b.data > a.data ? 1 : -1; });
    return lista.slice(0, 25);
  } catch (e) {
    return [];
  }
}

// ───────────────────────────────────────────────────────────
//  getLembretes — retorna lembretes ativos (fire_date <= hoje <= dataAlvo)
// ───────────────────────────────────────────────────────────
function getLembretes() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var ativos  = _todosLembretesPendentes(ss);
    var aniv    = _aniversariantesProximos(ss);
    var itens   = [];
    ativos.forEach(function(l) {
      itens.push({ tipo: 'lembrete', titulo: l.titulo, empresa: l.empresa, dataAlvo: l.dataAlvo, obs: l.obs });
    });
    aniv.forEach(function(a) {
      itens.push({ tipo: 'aniversario', titulo: 'Aniversário — ' + a.nome, empresa: a.empresa, dataAlvo: a.dataAniv, obs: '' });
    });
    itens.sort(function(a, b) {
      return a.dataAlvo < b.dataAlvo ? -1 : a.dataAlvo > b.dataAlvo ? 1 : 0;
    });
    return { ok: true, itens: itens };
  } catch (e) {
    return { ok: false, erro: e.message, itens: [] };
  }
}

// ───────────────────────────────────────────────────────────
//  salvarLembrete
// ───────────────────────────────────────────────────────────
function salvarLembrete(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = _getOrCreateSheet(ss, 'Base_Lembretes', [
      'Data/Hora Criacao', 'Empresa', 'Titulo', 'Data Alvo', 'Antecedencia (dias)', 'Data Disparo', 'Obs', 'Status'
    ]);
    var agora     = new Date();
    var dataAlvo  = new Date(dados.dataAlvo + 'T00:00:00');
    var antec     = parseInt(dados.antecedencia, 10) || 0;
    var disparo   = new Date(dataAlvo);
    disparo.setDate(disparo.getDate() - antec);
    var fmtDisp   = Utilities.formatDate(disparo,  Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var fmtAlvo   = Utilities.formatDate(dataAlvo, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.appendRow([
      Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      dados.empresa  || '',
      dados.titulo   || '',
      fmtAlvo,
      antec,
      fmtDisp,
      dados.obs      || '',
      'ativo'
    ]);
    return { ok: true };
  } catch (e) {
    return { ok: false, erro: e.message };
  }
}

// ───────────────────────────────────────────────────────────
//  concluirLembrete — marca status como 'concluído'
// ───────────────────────────────────────────────────────────
function concluirLembrete(dados) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Base_Lembretes');
    if (!sheet) return { ok: false, erro: 'Aba não encontrada' };
    var rows = sheet.getDataRange().getValues();
    var tz   = Session.getScriptTimeZone();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[7] || '').toLowerCase() !== 'ativo') continue;
      var titulo   = String(r[2] || '').trim();
      var empresa  = String(r[1] || '').trim();
      var dataAlvo = r[3] instanceof Date
        ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd')
        : String(r[3]);
      if (titulo === (dados.titulo || '').trim() &&
          empresa === (dados.empresa || '').trim() &&
          dataAlvo === (dados.dataAlvo || '').trim()) {
        sheet.getRange(i + 1, 8).setValue('concluído');
        return { ok: true };
      }
    }
    return { ok: false, erro: 'Lembrete não encontrado' };
  } catch (e) {
    return { ok: false, erro: e.message };
  }
}

// ───────────────────────────────────────────────────────────
//  _lembretesAtivos — lê Base_Lembretes, retorna itens onde
//  fire_date <= hoje <= dataAlvo e status = 'ativo'
// ───────────────────────────────────────────────────────────
function _lembretesAtivos(ss) {
  var sheet = ss.getSheetByName('Base_Lembretes');
  if (!sheet) return [];
  var hoje  = new Date(); hoje.setHours(0,0,0,0);
  var rows  = sheet.getDataRange().getValues();
  var ativos = [];
  for (var i = 1; i < rows.length; i++) {
    var r       = rows[i];
    var status  = String(r[7] || '').toLowerCase();
    if (status !== 'ativo') continue;
    var tz2 = Session.getScriptTimeZone();
    var dataAlvoStr = r[3] instanceof Date ? Utilities.formatDate(r[3], tz2, 'yyyy-MM-dd') : String(r[3]);
    var dataDispStr = r[5] instanceof Date ? Utilities.formatDate(r[5], tz2, 'yyyy-MM-dd') : String(r[5]);
    var dataAlvo  = new Date(dataAlvoStr + 'T00:00:00');
    var dataDisp  = new Date(dataDispStr + 'T00:00:00');
    if (isNaN(dataAlvo) || isNaN(dataDisp)) continue;
    if (hoje >= dataDisp && hoje <= dataAlvo) {
      ativos.push({
        empresa: String(r[1] || ''),
        titulo:  String(r[2] || ''),
        dataAlvo: String(r[3]),
        obs:     String(r[6] || '')
      });
    }
  }
  return ativos;
}

// ───────────────────────────────────────────────────────────
//  _todosLembretesPendentes — todos os lembretes com status ativo,
//  independente da data de disparo (para pontos no calendário)
// ───────────────────────────────────────────────────────────
function _todosLembretesPendentes(ss) {
  var sheet = ss.getSheetByName('Base_Lembretes');
  if (!sheet) return [];
  var tz   = Session.getScriptTimeZone();
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[7] || '').toLowerCase() !== 'ativo') continue;
    if (!r[3]) continue;
    var dataAlvo = r[3] instanceof Date
      ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd')
      : String(r[3]);
    result.push({
      empresa:  String(r[1] || ''),
      titulo:   String(r[2] || ''),
      dataAlvo: dataAlvo,
      obs:      String(r[6] || '')
    });
  }
  return result;
}

// ───────────────────────────────────────────────────────────
//  _aniversariantesProximos — colaboradores com aniversário
//  nos próximos 7 dias (inclusive hoje)
// ───────────────────────────────────────────────────────────
function _aniversariantesProximos(ss) {
  var sheet = ss.getSheetByName('Base_Colaboradores');
  if (!sheet) return [];
  var hoje  = new Date(); hoje.setHours(0,0,0,0);
  var rows  = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var nome      = String(r[2] || '').trim(); // col 3 = índice 2 (Nome)
    var empresa   = String(r[1] || '').trim(); // col 2 = índice 1 (Empresa)
    var nascStr   = r[5]; // col 6 = índice 5 (Nascimento)
    if (!nome || !nascStr) continue;
    var nasc = new Date(nascStr);
    if (isNaN(nasc)) continue;
    // próximo aniversário no ano corrente ou seguinte
    var anivEsteAno = new Date(hoje.getFullYear(), nasc.getMonth(), nasc.getDate());
    var aniv = anivEsteAno < hoje ? new Date(hoje.getFullYear() + 1, nasc.getMonth(), nasc.getDate()) : anivEsteAno;
    var diff = Math.round((aniv - hoje) / 86400000);
    if (diff >= 0 && diff <= 7) {
      var fmtAniv = Utilities.formatDate(aniv, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      result.push({ nome: nome, empresa: empresa, dataAniv: fmtAniv });
    }
  }
  return result;
}

function _gerarDocBasico(dados, agora) {
  var pastaEmpresa = _getOrCreateFolder(CONFIG.DRIVE_ROOT_ID, dados.empresa);
  var pastaFunc    = _getOrCreateFolder(pastaEmpresa.getId(), dados.nome);
  var pastaAdmBas  = _getOrCreateFolder(pastaFunc.getId(), 'Admissão');
  var dataFmt      = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var doc          = DocumentApp.create('Admissao Fisica - ' + dados.nome);
  var body         = doc.getBody();
  var estilo       = {};
  estilo[DocumentApp.Attribute.FONT_FAMILY]  = 'Arial';
  estilo[DocumentApp.Attribute.FONT_SIZE]    = 11;
  estilo[DocumentApp.Attribute.LINE_SPACING] = 1.5;

  body.setPageWidth(595.28).setPageHeight(841.89);
  body.appendParagraph('CRESCER CONSULTORIA').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('TERMO DE ADMISSAO').setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Empresa Cliente: ' + dados.empresa).setAttributes(estilo);
  body.appendParagraph('Nome: ' + dados.nome).setAttributes(estilo);
  body.appendParagraph('CPF: ' + dados.cpf).setAttributes(estilo);
  body.appendParagraph('Cargo: ' + (dados.cargo || '-')).setAttributes(estilo);
  body.appendParagraph('Data de Nascimento: ' + (dados.nascimento || '-')).setAttributes(estilo);
  body.appendParagraph('Telefone: ' + (dados.telefone || '-')).setAttributes(estilo);
  body.appendParagraph('Data de Admissao: ' + dataFmt).setAttributes(estilo);
  body.appendParagraph('\nDeclaro estar ciente e de acordo com as normas internas da empresa, incluindo Regimento Interno, Banco de Horas e Politica de Privacidade (LGPD).').setAttributes(estilo);
  body.appendParagraph('Local e Data: __________________, ' + dataFmt).setAttributes(estilo);
  body.appendParagraph('\n\n__________________________________');
  body.appendParagraph('Assinatura - ' + dados.nome);
  body.appendParagraph('\n\n__________________________________');
  body.appendParagraph('Assinatura - Responsavel RH');
  doc.saveAndClose();

  var docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(pastaAdmBas);
  return docFile.getUrl();
}
