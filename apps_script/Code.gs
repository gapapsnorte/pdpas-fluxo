/**
 * Ferramenta PDPAS (versão web) - ponte com o Google Drive e a Planilha mestra.
 *
 * PORTÁTIL: não tem nome de conta, código nem ID escrito aqui. Tudo é criado na conta que rodar `preparar`
 * e guardado em Propriedades do script. Para mudar de conta, veja TRANSPLANTAR.md.
 *
 * O que faz:
 *  - guarda o backup dos pedidos (um arquivo .json na pasta "Ferramenta PDPAS" do Drive);
 *  - mantém a "Planilha mestra" com uma linha por pedido (não apaga linhas: pedido apagado fica como "apagado").
 *
 * Segurança: toda chamada precisa do CODIGO (Propriedades do script). Sem ele, nada é lido nem gravado.
 */
const NOME_PASTA = 'Ferramenta PDPAS';
const NOME_PLANILHA = 'PDPAS - Planilha mestra (histórico)';
const NOME_BACKUP = 'backup_pdpas.json';
const ABA = 'Pedidos';
const COLUNAS = ['id', 'atualizado', 'criado', 'codigo_ses', 'item', 'tipo', 'quantidade_total', 'valor_estimado', 'unidades',
  'fase', 'passos_feitos', 'passos_total', 'numeros_sei', 'minutas', 'situacao'];

function props_() { return PropertiesService.getScriptProperties(); }

/** Rode UMA vez (menu Executar > preparar). Cria a pasta e a planilha mestra na conta que está rodando. */
function preparar() {
  const p = props_();
  let pasta = p.getProperty('PASTA_ID') ? DriveApp.getFolderById(p.getProperty('PASTA_ID')) : null;
  if (!pasta) {
    const it = DriveApp.getFoldersByName(NOME_PASTA);
    pasta = it.hasNext() ? it.next() : DriveApp.createFolder(NOME_PASTA);
    p.setProperty('PASTA_ID', pasta.getId());
  }
  let ss = null;
  if (p.getProperty('PLANILHA_ID')) { try { ss = SpreadsheetApp.openById(p.getProperty('PLANILHA_ID')); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(NOME_PLANILHA);
    DriveApp.getFileById(ss.getId()).moveTo(pasta);
    p.setProperty('PLANILHA_ID', ss.getId());
  }
  let aba = ss.getSheetByName(ABA);
  if (!aba) { aba = ss.getSheets()[0]; aba.setName(ABA); }
  if (aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, COLUNAS.length).setValues([COLUNAS]).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  Logger.log('Pronto. Pasta: ' + pasta.getUrl() + ' | Planilha: ' + ss.getUrl());
  if (!p.getProperty('CODIGO')) {
    Logger.log('FALTA definir o CODIGO: Configurações do projeto > Propriedades do script > adicionar CODIGO = uma senha longa que só você sabe.');
  }
}

function resp_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function conferir_(codigo) {
  const certo = props_().getProperty('CODIGO');
  if (!certo || String(certo).length < 8) return false;          // sem código forte definido, recusa tudo
  const ok = typeof codigo === 'string' && codigo.length === certo.length && codigo === certo;
  if (!ok) Utilities.sleep(1000);                                 // atrasa tentativas de adivinhar
  return ok;
}

function doGet() { return resp_({ ok: true, servico: 'Ferramenta PDPAS', preparado: !!props_().getProperty('PLANILHA_ID') }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    if (!conferir_(d.codigo)) return resp_({ ok: false, erro: 'Código incorreto ou não definido no script.' });
    lock.waitLock(20000);
    switch (d.acao) {
      case 'ping': return resp_({ ok: true });
      case 'salvar_backup': return resp_(salvarBackup_(String(d.conteudo || '')));
      case 'ler_backup': return resp_(lerBackup_());
      case 'linha': return resp_(linha_(d.linha));
      case 'linhas': return resp_(linhas_(d.linhas || []));
      case 'apagado': return resp_(apagado_(String(d.id || '')));
      default: return resp_({ ok: false, erro: 'ação desconhecida' });
    }
  } catch (err) {
    return resp_({ ok: false, erro: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* sem lock */ }
  }
}

function pasta_() {
  const id = props_().getProperty('PASTA_ID');
  if (!id) throw new Error('Rode a função preparar antes.');
  return DriveApp.getFolderById(id);
}

function salvarBackup_(conteudo) {
  if (conteudo.length < 20) throw new Error('backup vazio');
  const pasta = pasta_(), it = pasta.getFilesByName(NOME_BACKUP);
  if (it.hasNext()) { it.next().setContent(conteudo); } else { pasta.createFile(NOME_BACKUP, conteudo, MimeType.PLAIN_TEXT); }
  return { ok: true, quando: new Date().toISOString() };
}

function lerBackup_() {
  const it = pasta_().getFilesByName(NOME_BACKUP);
  if (!it.hasNext()) return { ok: true, conteudo: '' };
  return { ok: true, conteudo: it.next().getBlob().getDataAsString() };
}

function aba_() {
  const id = props_().getProperty('PLANILHA_ID');
  if (!id) throw new Error('Rode a função preparar antes.');
  return SpreadsheetApp.openById(id).getSheetByName(ABA);
}

function valores_(linha) { return COLUNAS.map(function (c) { var v = linha[c]; return v === undefined || v === null ? '' : v; }); }

function linha_(linha) {
  if (!linha || !linha.id) throw new Error('linha sem id');
  const aba = aba_(), ult = aba.getLastRow();
  const ids = ult > 1 ? aba.getRange(2, 1, ult - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const i = ids.indexOf(String(linha.id));
  if (i >= 0) { aba.getRange(i + 2, 1, 1, COLUNAS.length).setValues([valores_(linha)]); }
  else { aba.appendRow(valores_(linha)); }
  return { ok: true };
}

function linhas_(lista) {
  lista.forEach(function (l) { linha_(l); });
  return { ok: true, n: lista.length };
}

function apagado_(id) {
  const aba = aba_(), ult = aba.getLastRow();
  if (ult < 2) return { ok: true };
  const ids = aba.getRange(2, 1, ult - 1, 1).getValues().map(function (r) { return String(r[0]); });
  const i = ids.indexOf(id);
  if (i >= 0) { aba.getRange(i + 2, COLUNAS.indexOf('situacao') + 1).setValue('apagado em ' + new Date().toLocaleDateString('pt-BR')); }
  return { ok: true };
}
