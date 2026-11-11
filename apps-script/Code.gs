/**
 * GM CelebrArt — Central (Google Apps Script)
 *
 * Instalação (resumo — ver LEIA-ME.md):
 *  1. Crie uma Planilha Google → Extensões → Apps Script.
 *  2. Cole este arquivo e os arquivos .html (mesmos nomes, sem extensão).
 *  3. Rode a função setup() uma vez (autorize).
 *  4. Suba as imagens na pasta "GM CelebrArt/Imagens" criada no Drive.
 *  5. Implantar → Nova implantação → App da Web.
 */

const ABAS = {
  Config:     ['Chave', 'Valor', 'Descrição'],
  Categorias: ['Nome', 'Oculta do cliente?'],
  Acervo:     ['Nome', 'Custo (R$)', 'Categoria'],
  Kits:       ['Kit', 'Preço (R$)'],
  ItensKit:   ['Item no checklist', 'Nome no contrato/proposta', 'Bronze', 'Prata', 'Ouro'],
  Extras:     ['Nome', 'Descrição', 'Preço (R$)', 'Item no checklist'],
  Temas:      ['Tema', 'Emoji'],
  Modelos:    ['ID', 'Nome', 'Criado em', 'Dados (JSON)'],
  Propostas:  ['Número', 'Emitida em', 'Cliente', 'WhatsApp', 'Tema', 'Data do evento', 'Modalidade',
               'Kit', 'Total', 'Sinal', 'Restante', 'Arquivo', 'Atualizada em', 'Por'],
  Contratos:  ['Gerado em', 'Documento', 'Cliente', 'CPF/CNPJ', 'Telefone', 'Endereço', 'Bairro', 'Tema',
               'Kit', 'Modalidade', 'Data do evento', 'Montagem/Entrega', 'Desmontagem/Devolução',
               'Total', 'Sinal', 'Restante', 'Por']
};

/* ═════════════ WEB APP ═════════════ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('GM CelebrArt — Central')
    // no Apps Script o <meta viewport> do HTML é ignorado: precisa vir daqui
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* API para a página hospedada fora do Google (GitHub Pages).
   A página manda POST text/plain {pin, quem, nome, args} e recebe {ok, v} ou {ok:false, e}. */
function doPost(e) {
  let saida;
  try {
    const b = JSON.parse(e.postData.contents);
    saida = { ok: true, v: api(b.pin, b.quem, b.nome, JSON.stringify(b.args || [])) };
  } catch (err) {
    saida = { ok: false, e: String(err && err.message || err) };
  }
  if (saida.ok && saida.v === undefined) saida.v = null;
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}


/* ═════════════ PORTÃO (PIN) ═════════════
   O app roda como o DONO, então nenhuma função pública pode devolver dados sem o PIN.
   Funções terminadas em "_" não são chamáveis pelo navegador; só api() é. */

const API_ = {
  getBoot: getBoot_, getImagens: getImagens_, salvarImagem: salvarImagem_, removerImagem: removerImagem_,
  acervoListar: acervoListar_, acervoSalvar: acervoSalvar_, acervoRemover: acervoRemover_,
  modelosListar: modelosListar_, modeloSalvar: modeloSalvar_, modeloApagar: modeloApagar_,
  propostaReservarNumero: propostaReservarNumero_, salvarArquivoProposta: salvarArquivoProposta_,
  propostaRegistrar: propostaRegistrar_, contratoRegistrar: contratoRegistrar_
};
let QUEM_ = '';

function api(pin, quem, nome, argsJson) {
  const cache = CacheService.getScriptCache();
  const falhas = Number(cache.get('pin_falhas') || 0);
  if (falhas >= 8) throw new Error('PIN_BLOQUEADO: muitas tentativas erradas. Aguarde 10 minutos.');
  const certo = txt_(lerConfig_().pin);
  if (!certo) throw new Error('PIN_NAO_DEFINIDO: rode criarPin() no editor do Apps Script.');
  if (txt_(pin) !== certo) {
    cache.put('pin_falhas', String(falhas + 1), 600);
    throw new Error('PIN_INVALIDO');
  }
  if (falhas) cache.remove('pin_falhas');
  if (!Object.prototype.hasOwnProperty.call(API_, nome)) throw new Error('Função desconhecida.');
  QUEM_ = txt_(quem).slice(0, 40);
  return API_[nome].apply(null, JSON.parse(argsJson || '[]'));
}

/** Rodar 1x no editor se a planilha já existia antes do PIN (setup() novo já cria). Mostra o PIN no Registro. */
function criarPin() {
  const sh = aba_('Config');
  if (txt_(lerConfig_().pin)) { Logger.log('Já existe PIN. Para trocar, edite a aba Config (linha "pin").'); return; }
  const pin = String(Math.floor(100000 + Math.random() * 900000));
  gravarConfig_('pin', "'" + pin);
  Logger.log('PIN criado: ' + pin + '  (fica na aba Config, linha "pin")');
}

/* ═════════════ LEITURA INICIAL ═════════════ */

function getBoot_() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('Config')) {
    throw new Error('Planilha não configurada: rode a função setup() uma vez no editor do Apps Script.');
  }
  const cats = lerTabela_('Categorias')
    .map(r => ({ id: txt_(r.nome), nome: txt_(r.nome), oculta: sim_(r.oculta_do_cliente) }))
    .filter(c => c.id);
  const kits = lerTabela_('Kits')
    .map(r => ({ nome: txt_(r.kit), preco: r.preco === '' ? '' : num_(r.preco) }))
    .filter(k => k.nome);
  const extras = lerTabela_('Extras')
    .map(r => ({ nome: txt_(r.nome), desc: txt_(r.descricao), preco: num_(r.preco), chk: txt_(r.item_no_checklist) }))
    .filter(e => e.nome);
  const temas = lerTabela_('Temas')
    .map(r => ({ tema: txt_(r.tema), emoji: txt_(r.emoji) }))
    .filter(t => t.tema);

  // ItensKit: a partir da 3ª coluna, cada cabeçalho é um kit (dá para criar kits novos só acrescentando colunas)
  const v = aba_('ItensKit').getDataRange().getValues();
  const kitCols = v[0].slice(2).map(txt_).filter(Boolean);
  const itensKit = v.slice(1)
    .filter(r => txt_(r[0]))
    .map(r => {
      const qtd = {};
      kitCols.forEach((k, i) => { qtd[k] = num_(r[i + 2]); });
      return { item: txt_(r[0]), contrato: txt_(r[1]), qtd: qtd };
    });

  return {
    cfg: (function () { const c = lerConfig_(); delete c.pin; delete c.pasta_id; delete c.seq_mes; delete c.seq_num; return c; })(),
    cats: cats,
    acervo: lerAcervo_(),
    kits: kits,
    kitCols: kitCols,
    itensKit: itensKit,
    extras: extras,
    temas: temas,
    planilhaUrl: ss.getUrl()
  };
}

/* Imagens ficam no Drive (pasta "Imagens"). Nomes:
   logo.*            → logo da proposta
   logo-contrato.*   → logo do contrato/checklist
   kit - <Tema> - <Kit>.jpg → foto da montagem no checklist */
function getImagens_() {
  const out = { logo: null, logoContrato: null, fotos: {} };
  const it = pasta_('Imagens').getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.isTrashed()) continue;
    const mime = f.getMimeType();
    if (!/^image\//.test(mime)) continue;
    const base = f.getName().replace(/\.[^.]+$/, '').trim();
    const url = 'data:' + mime + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes());
    if (/^logo$/i.test(base)) out.logo = url;
    else if (/^logo-contrato$/i.test(base)) out.logoContrato = url;
    else {
      const m = base.match(/^kit\s*-\s*(.+?)\s*-\s*(.+)$/i);
      if (m) {
        const tema = m[1].trim().toLowerCase(), kit = m[2].trim().toLowerCase();
        (out.fotos[tema] = out.fotos[tema] || {})[kit] = url;
      }
    }
  }
  return out;
}

function salvarImagem_(nomeBase, dataUrl) {
  const m = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('Imagem inválida.');
  const pasta = pasta_('Imagens');
  removerImagem_(nomeBase);
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg').replace('svg+xml', 'svg');
  pasta.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nomeBase + '.' + ext));
  return true;
}

function removerImagem_(nomeBase) {
  const it = pasta_('Imagens').getFiles();
  const re = new RegExp('^' + nomeBase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\.[^.]+$', 'i');
  while (it.hasNext()) {
    const f = it.next();
    if (re.test(f.getName())) f.setTrashed(true);
  }
  return true;
}

/* ═════════════ ACERVO ═════════════ */

function lerAcervo_() {
  return lerTabela_('Acervo')
    .map(r => ({ n: txt_(r.nome), c: num_(r.custo), cat: txt_(r.categoria) }))
    .filter(x => x.n);
}

function acervoListar_() { return lerAcervo_(); }

/** nomeAntigo vazio = item novo. Devolve o acervo atualizado (pega também o que o outro aparelho mudou). */
function acervoSalvar_(nomeAntigo, item) {
  return comLock_(() => {
    const nome = txt_(item.n), custo = Math.round(num_(item.c) * 100) / 100, cat = txt_(item.cat);
    if (!nome) throw new Error('Nome vazio.');
    if (!(custo >= 0)) throw new Error('Valor inválido.');
    const sh = aba_('Acervo');
    const nomes = colunaA_(sh);
    const iAntigo = nomeAntigo ? nomes.indexOf(txt_(nomeAntigo)) : -1;
    const iNovo = nomes.indexOf(nome);
    if (iNovo >= 0 && iNovo !== iAntigo) throw new Error('Já existe um item chamado "' + nome + '".');
    if (iAntigo >= 0) sh.getRange(iAntigo + 2, 1, 1, 3).setValues([[nome, custo, cat]]);
    else sh.appendRow([nome, custo, cat]);
    return lerAcervo_();
  });
}

function acervoRemover_(nome) {
  return comLock_(() => {
    const sh = aba_('Acervo');
    const i = colunaA_(sh).indexOf(txt_(nome));
    if (i >= 0) sh.deleteRow(i + 2);
    return lerAcervo_();
  });
}

/* ═════════════ MODELOS DE PROPOSTA ═════════════ */

function modelosListar_() {
  return lerTabela_('Modelos').map(r => {
    let dados = null;
    try { dados = JSON.parse(r.dados); } catch (e) {}
    return { id: txt_(r.id), nome: txt_(r.nome), em: dataIso_(r.criado_em), dados: dados };
  }).filter(m => m.id && m.dados).reverse();   // mais recentes primeiro
}

function modeloSalvar_(nome, dados) {
  const json = JSON.stringify(dados);
  if (json.length > 49000) throw new Error('Modelo grande demais para a planilha.');
  return comLock_(() => {
    aba_('Modelos').appendRow(['m' + Date.now(), txt_(nome) || 'Modelo', hoje_(), json]);
    return modelosListar_();
  });
}

function modeloApagar_(id) {
  return comLock_(() => {
    const sh = aba_('Modelos');
    const i = colunaA_(sh).indexOf(txt_(id));
    if (i >= 0) sh.deleteRow(i + 2);
    return modelosListar_();
  });
}

/* ═════════════ PROPOSTAS ═════════════ */

/** Número AAMM-NN, único entre aparelhos, reinicia todo mês. */
function propostaReservarNumero_() {
  return comLock_(() => {
    const aamm = Utilities.formatDate(new Date(), fuso_(), 'yyMM');
    const cfg = lerConfig_();
    let seq = String(cfg.seq_mes) === aamm ? (Number(cfg.seq_num) || 0) : 0;
    seq++;
    gravarConfig_('seq_mes', "'" + aamm);
    gravarConfig_('seq_num', seq);
    return aamm + '-' + String(seq).padStart(2, '0');
  });
}

/** Salva PDF/PNG da proposta no Drive. base64 sem o prefixo data:. */
function salvarArquivoProposta_(nomeArquivo, base64, mime) {
  const pasta = pasta_('Propostas');
  const antigos = pasta.getFilesByName(nomeArquivo);
  while (antigos.hasNext()) antigos.next().setTrashed(true);
  const f = pasta.createFile(Utilities.newBlob(Utilities.base64Decode(base64), mime, nomeArquivo));
  return { id: f.getId(), url: f.getUrl() };
}

/** Grava/atualiza a linha da proposta (chave = número). */
function propostaRegistrar_(p) {
  return comLock_(() => {
    const sh = aba_('Propostas');
    const linha = [
      "'" + txt_(p.numero), p.emitida || hoje_(), txt_(p.cliente), txt_(p.zap), txt_(p.tema),
      txt_(p.data), txt_(p.modalidade), txt_(p.kit), num_(p.total), num_(p.sinal), num_(p.restante),
      txt_(p.arquivo), agora_(), usuario_()
    ];
    const i = p.numero ? colunaA_(sh).indexOf(txt_(p.numero)) : -1;
    if (i >= 0) {
      if (!p.arquivo) linha[11] = sh.getRange(i + 2, 12).getValue();   // não apaga o link anterior
      sh.getRange(i + 2, 1, 1, linha.length).setValues([linha]);
    } else sh.appendRow(linha);
    return true;
  });
}

/* ═════════════ CONTRATOS ═════════════ */

function contratoRegistrar_(c) {
  return comLock_(() => {
    aba_('Contratos').appendRow([
      agora_(), txt_(c.documento), txt_(c.cliente), "'" + txt_(c.cpf), txt_(c.tel), txt_(c.end),
      txt_(c.bairro), txt_(c.tema), txt_(c.kit), txt_(c.modalidade), txt_(c.data), txt_(c.inicio),
      txt_(c.fim), num_(c.total), num_(c.sinal), num_(c.restante), usuario_()
    ]);
    return true;
  });
}

/* ═════════════ SETUP (rodar 1x no editor) ═════════════ */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  criarAba_('Config', [
    ['empresa', 'GM Celebrart', 'Nome no contrato'],
    ['endereco', 'Rua Cyro Correia Pereira', 'Endereço da contratada'],
    ['numero', '3689', ''],
    ['cidade', 'Curitiba', 'Cidade da assinatura do contrato'],
    ['regiao', 'Curitiba e região', 'Rodapé da proposta'],
    ['contato1_nome', 'Gabriel', ''],
    ['contato1_tel', '(41) 99781-1287', ''],
    ['contato2_nome', 'Mayara', ''],
    ['contato2_tel', '(41) 99820-4206', ''],
    ['instagram', '@gmcelebrart', ''],
    ['pct_padrao', 20, '% sobre o custo de compra (aba Orçamento)'],
    ['sinal_padrao', 50, 'Sinal padrão em R$'],
    ['validade_padrao', '7 dias', 'Validade padrão da proposta'],
    ['multa_atraso', 20, 'R$ por dia de atraso (Pegue e Monte)'],
    ['seq_mes', '', 'Controle interno da numeração — não editar'],
    ['seq_num', '', 'Controle interno da numeração — não editar'],
    ['pin', "'" + String(Math.floor(100000 + Math.random() * 900000)), 'PIN de acesso ao app — troque quando quiser'],
    ['pasta_id', '', 'Pasta do Drive (preenchida pelo setup)']
  ]);

  criarAba_('Categorias', [
    ['Materiais de uso e consumo', 'sim'], ['Painéis', ''], ['Mesas e estantes', ''], ['Cilindros', ''],
    ['Boleiras e bandejas', ''], ['Capas e tecidos', ''], ['Displays', ''], ['Flores', ''],
    ['Decoração', ''], ['Tapetes', ''], ['Bolos fake', '']
  ]);

  const novaAcervo = criarAba_('Acervo', ACERVO_INICIAL_());
  if (novaAcervo) {
    const sh = aba_('Acervo');
    sh.getRange('B2:B').setNumberFormat('"R$" #,##0.00');
    sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(aba_('Categorias').getRange('A2:A'), true).setAllowInvalid(true).build());
  }

  criarAba_('Kits', [
    ['Bronze', 120], ['Prata', 155], ['Ouro', 180], ['Chá de Bebê', ''], ['Adulto', ''], ['Personalizado', '']
  ]);

  // Corrige o bug do original: o Ouro estava sem o Número em MDF no checklist.
  criarAba_('ItensKit', [
    ['Painel redondo (6 peças)',     'Painel redondo',           1, 1, 1],
    ['Trio de cilindros (3 peças)',  'Trio de cilindros',        1, 1, 1],
    ['Capas no tema',                'Capas no tema',            4, 4, 4],
    ['Vasos decorativos',            'Vasos com arbusto',        0, 2, 2],
    ['Porta bolo',                   'Porta-bolo',               1, 1, 1],
    ['Porta docinhos',               'Porta-docinhos',           4, 4, 4],
    ['Totens de mesa',               'Totem de mesa',            0, 5, 5],
    ['Totens de chão',               '',                         0, 0, 0],
    ['Tapete decorativo',            'Tapete decorativo',        1, 1, 1],
    ['Cômoda Branca (2 peças)',      'Cômoda branca',            0, 0, 1],
    ['Painel Romano',                '',                         0, 0, 0],
    ['Número em MDF',                'Número em MDF',            0, 1, 1],
    ['Estante de lembrancinhas',     'Estante de lembrancinhas', 0, 0, 1],
    ['Arbusto de chão',              '',                         0, 0, 0],
    ['Displays de chão',             '',                         0, 0, 0],
    ['Cômoda Fake',                  '',                         0, 0, 0],
    ['Bolo Fake',                    '',                         0, 0, 0]
  ]);

  criarAba_('Extras', [
    ['Estante de lembrancinhas', 'Branca ou cru',     20, 'Estante de lembrancinhas'],
    ['Número em MDF',            'Do 0 ao 9 (cada)',  15, 'Número em MDF'],
    ['2 Displays de chão',       'Safari Baby',       25, 'Displays de chão'],
    ['Cômoda fake',              'Branca decorativa', 35, 'Cômoda Fake'],
    ['Bolo fake',                'Safari Baby',       20, 'Bolo Fake']
  ]);

  criarAba_('Temas', [
    ['Safari Baby', '🦁'], ['Patrulha Canina', '🐾'], ['Hello Kitty', '🎀'], ['Chá de Bebê', '🍼']
  ]);

  criarAba_('Modelos', []);
  criarAba_('Propostas', []);
  criarAba_('Contratos', []);

  // aba padrão "Página1"/"Sheet1" vazia sai
  ss.getSheets().forEach(s => {
    if (!ABAS[s.getName()] && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  const raiz = pastaRaiz_();
  pasta_('Imagens'); pasta_('Propostas');
  Logger.log('PIN de acesso: ' + lerConfig_().pin + '  (aba Config, linha "pin")');
  Logger.log('Pronto! Suba as imagens em: ' + pasta_('Imagens').getUrl());
  Logger.log('Pasta principal (compartilhe com quem vai usar): ' + raiz.getUrl());
}

/* ═════════════ HELPERS ═════════════ */

function aba_(nome) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não existe. Rode setup().');
  return sh;
}

/** Cria a aba com cabeçalho e dados iniciais só se ela ainda não existir. Devolve true se criou. */
function criarAba_(nome, linhas) {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(nome)) return false;
  const sh = ss.insertSheet(nome);
  const cab = ABAS[nome];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#fdf3f6');
  sh.setFrozenRows(1);
  if (linhas.length) sh.getRange(2, 1, linhas.length, linhas[0].length).setValues(linhas);
  sh.autoResizeColumns(1, cab.length);
  return true;
}

function chave_(h) {
  return String(h).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\(.*?\)/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

function lerTabela_(nome) {
  const v = aba_(nome).getDataRange().getValues();
  if (v.length < 2) return [];
  const cab = v[0].map(chave_);
  return v.slice(1).map(r => {
    const o = {};
    cab.forEach((k, i) => { o[k] = r[i] instanceof Date ? dataIso_(r[i]) : r[i]; });
    return o;
  });
}

function colunaA_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => txt_(r[0])) : [];
}

function lerConfig_() {
  const o = {};
  lerTabela_('Config').forEach(r => { if (txt_(r.chave)) o[txt_(r.chave)] = r.valor; });
  return o;
}

function gravarConfig_(chave, valor) {
  const sh = aba_('Config');
  const i = colunaA_(sh).indexOf(chave);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(valor);
  else sh.appendRow([chave, valor, '']);
}

function pastaRaiz_() {
  const id = lerConfig_().pasta_id;
  if (id) { try { return DriveApp.getFolderById(String(id)); } catch (e) {} }
  const p = DriveApp.createFolder('GM CelebrArt');
  gravarConfig_('pasta_id', p.getId());
  return p;
}

function pasta_(nome) {
  const raiz = pastaRaiz_();
  const it = raiz.getFoldersByName(nome);
  return it.hasNext() ? it.next() : raiz.createFolder(nome);
}

function comLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function txt_(v) { return v == null ? '' : String(v).trim(); }
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v || '').replace(/[R$\s]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');   // formato 1.234,56
  const n = Number(s);
  return isFinite(n) ? n : 0;
}
function sim_(v) { return v === true || /^(s|sim|x|true|1)$/i.test(txt_(v)); }
function fuso_() { return Session.getScriptTimeZone() || 'America/Sao_Paulo'; }
function hoje_() { return Utilities.formatDate(new Date(), fuso_(), 'yyyy-MM-dd'); }
function agora_() { return Utilities.formatDate(new Date(), fuso_(), 'yyyy-MM-dd HH:mm'); }
function dataIso_(d) { return d instanceof Date ? Utilities.formatDate(d, fuso_(), 'yyyy-MM-dd') : txt_(d); }
function usuario_() { return QUEM_ || 'sem nome'; }

function ACERVO_INICIAL_() {
  return [
    ['Spray de balão', 60, 'Materiais de uso e consumo'],
    ['Presilhas', 25.25, 'Materiais de uso e consumo'],
    ['Elástico amarelo', 10.9, 'Materiais de uso e consumo'],
    ['Painel romano ripado', 184.2, 'Painéis'],
    ['Painel romano metálico', 125, 'Painéis'],
    ['Painel redondo metálico', 125, 'Painéis'],
    ['Mesas treliças', 253.01, 'Mesas e estantes'],
    ['Cômoda fake amarela', 180, 'Mesas e estantes'],
    ['Cômoda fake branca', 90, 'Mesas e estantes'],
    ['Trio de cilindro ripado', 442.2, 'Cilindros'],
    ['Trio de cilindro', 121.42, 'Cilindros'],
    ['Kit boleiras espelhadas', 174, 'Boleiras e bandejas'],
    ['Kit boleiras bambu (5)', 91.22, 'Boleiras e bandejas'],
    ['Kit boleiras cristal', 84.07, 'Boleiras e bandejas'],
    ['Vasos grandes bege (2)', 70.92, 'Boleiras e bandejas'],
    ['Kit boleiras rosa candy', 68.2, 'Boleiras e bandejas'],
    ['Kit boleiras Patrulha Canina', 67.79, 'Boleiras e bandejas'],
    ['Boleira amarelo candy', 67.62, 'Boleiras e bandejas'],
    ['Kit boleiras dourado', 66.68, 'Boleiras e bandejas'],
    ['Bandejas flor rosa candy (2)', 53.48, 'Boleiras e bandejas'],
    ['Bandeja bambu quadrada', 48.99, 'Boleiras e bandejas'],
    ['Kit 2 mesinhas de mesa bege', 44.25, 'Boleiras e bandejas'],
    ['Bandeja bambu comprida', 41.61, 'Boleiras e bandejas'],
    ['Kit 2 mesinhas de mesa azul bebê', 37.37, 'Boleiras e bandejas'],
    ['Bandejas desconstruídas bege (3)', 36.84, 'Boleiras e bandejas'],
    ['Kit 2 bandejas ovais bambu', 36.42, 'Boleiras e bandejas'],
    ['Kit 2 bandejas lilás', 24.76, 'Boleiras e bandejas'],
    ['Kit 2 bandejas laranja', 24.76, 'Boleiras e bandejas'],
    ['Bandeja bambu redonda', 22.18, 'Boleiras e bandejas'],
    ['Vasos grandes candy (2)', 21.53, 'Boleiras e bandejas'],
    ['Vasos de cristal (2)', 29.7, 'Boleiras e bandejas'],
    ['Capas cilindro + painel redondo Safari Baby', 119.51, 'Capas e tecidos'],
    ['Capas cilindro + painel redondo Patrulha Canina', 108.72, 'Capas e tecidos'],
    ['Capas cilindro + painel redondo Guerreiras K-pop', 105.63, 'Capas e tecidos'],
    ['Tecido de linho', 89.7, 'Capas e tecidos'],
    ['Capas cilindro + painel redondo branca', 80.73, 'Capas e tecidos'],
    ['Capas cilindro + painel romano Jardim Encantado', 79.83, 'Capas e tecidos'],
    ['Capas cilindro + painel redondo Dino Baby', 76.99, 'Capas e tecidos'],
    ['Capa bege e marrom painel romano', 66.59, 'Capas e tecidos'],
    ['Capa painel romano azul bebê', 52.87, 'Capas e tecidos'],
    ['Capa painel romano rosa bebê', 52.87, 'Capas e tecidos'],
    ['Capa cilindro rosa', 42.98, 'Capas e tecidos'],
    ['Número MDF', 230, 'Displays'],
    ['Placa Divino Espírito Santo', 165, 'Displays'],
    ['Display de mesa Safari Baby', 83.98, 'Displays'],
    ['Kit display de mesa flores, borboleta e coração', 82.73, 'Displays'],
    ['Display de chão borboletas rosa', 60, 'Displays'],
    ['Display de mesa Dino Baby', 47.95, 'Displays'],
    ['Display de mesa Patrulha Canina', 38.61, 'Displays'],
    ['Display de mesa Guerreiras K-pop', 37.8, 'Displays'],
    ['Display de chão Dino Baby', 0, 'Displays'],
    ['Buquês crisântemo & botão (2)', 113.92, 'Flores'],
    ['Buquê crisântemos e botões rosa e branco', 48, 'Flores'],
    ['Flor artificial branco e rosa', 43.09, 'Flores'],
    ['Mosquitinho', 39, 'Flores'],
    ['Kit 2 arbustos artificiais pequenos', 23.6, 'Flores'],
    ['Cubo Baby', 333.95, 'Decoração'],
    ['LED Happy Birthday', 186.65, 'Decoração'],
    ['Urso grande', 162.86, 'Decoração'],
    ['Escadinha lembrancinha branca', 46.82, 'Decoração'],
    ['Escadinha lembrancinha bege', 46.82, 'Decoração'],
    ['Vaso de chão bege grande', 45, 'Decoração'],
    ['Estátua de mesa anjo 10cm', 25, 'Decoração'],
    ['Estátua de mesa anjo ajoelhado 20cm', 20, 'Decoração'],
    ['Estátua de mesa anjo 14cm', 15, 'Decoração'],
    ['Urso pequeno', 42.41, 'Decoração'],
    ['Mini cavalete', 39.97, 'Decoração'],
    ['Base chão balão', 24.9, 'Decoração'],
    ['Tapete rosa', 103.8, 'Tapetes'],
    ['Tapete de crochê', 90, 'Tapetes'],
    ['Tapete bege', 73.4, 'Tapetes'],
    ['Tapete verde', 71.58, 'Tapetes'],
    ['Tapete azul royal', 70.81, 'Tapetes'],
    ['Bolo fake Guerreiras K-pop', 74.5, 'Bolos fake'],
    ['Bolo fake Safari Baby', 64.9, 'Bolos fake'],
    ['Estrela Pendente Glitter Dourado 26cm (9)', 261.9, 'Decoração'],
    ['Balões Estrela Verde e Vermelho', 29.8, 'Decoração'],
    ['Festão', 78, 'Decoração'],
    ['Guirlanda (8)', 472, 'Decoração'],
    ['Flor Dourada Enfeite Arvore (6)', 29.92, 'Decoração'],
    ['Flor Vermelha Enfeite Arvore (6)', 29.92, 'Decoração'],
    ['Saia Enfeite Arvore Vermelha e Verde', 39.92, 'Decoração'],
    ['Kit Bolas Dourada Pendentes Enfeite Arvore', 104, 'Decoração'],
    ['Kit Bola Vermelha Pendentes Enfeite Arvore', 104, 'Decoração'],
    ['Estrela Topo Arvore Dourada', 23.99, 'Decoração'],
    ['Pisca Pisca Led Branco Quente', 125, 'Decoração'],
    ['Arvore de Natal 3M', 789.9, 'Decoração'],
    ['Mini Velas a Pilha', 70, 'Decoração'],
  ];
}
