/**
 * ============================================================
 *  MERCADINHO DE TODOS — Motor do sistema (Google Apps Script)
 *  Cartão de TODOS · Campo Grande
 * ============================================================
 *  1) Troque a SENHA_ADMIN logo abaixo (senha da Área do gestor).
 *  2) Rode a função  setup  uma vez (botão ▶ Executar).
 *  3) (Opcional) Rode  cargaInicial  para cadastrar parceiros,
 *     itens e colaboradores listados no fim deste arquivo.
 *  4) Implantar > Nova implantação > App da Web
 *       Executar como: Eu  ·  Quem pode acessar: Qualquer pessoa
 *     Copie a URL gerada e cole no index.html (API_URL).
 * ============================================================
 */
const SENHA_ADMIN = 'TROQUE-ESTA-SENHA';
const TZ = 'America/Campo_Grande';
const SETORES = ['Vendas', 'Adimplência', 'Retenção', 'Administrativo', 'Parcerias', 'Gerência', 'Clínica', 'Higienização'];
const CATEGORIAS = { COLABORADOR: 'Colaborador', FILIADO: 'Filiado', AVULSA: 'Venda avulsa' };
const FORMAS = ['App Cartão de TODOS', 'Pix', 'Cartão', 'Dinheiro'];
// Unidades de negócio. Fiado (pagar depois) só existe para a unidade Cartão de TODOS.
const UN = { CDT: 'Cartão de TODOS', AMOR: 'AmorSaúde', ACOES: 'Ações Externas' };
const NOME_CONTA_ACOES = 'Ações Externas';
// true = cadastro feito pela própria pessoa fica aguardando aprovação na Área do gestor
const APROVAR_CADASTROS = true;

const ABAS = {
  Parceiros:     ['id', 'nome', 'contato', 'whatsapp', 'pagamento', 'ativo', 'criadoEm'],
  Itens:         ['id', 'parceiroId', 'nome', 'preco', 'estoque', 'estoqueMinimo', 'ativo', 'criadoEm'],
  Colaboradores: ['id', 'nome', 'setor', 'pinHash', 'ativo', 'criadoEm', 'unidade', 'whatsapp', 'funcao'],
  Retiradas:     ['id', 'pedidoId', 'dataHora', 'colaboradorId', 'colaborador', 'setor', 'parceiroId', 'parceiro',
                  'itemId', 'item', 'qtd', 'precoUnit', 'total', 'fotoId', 'modo', 'status', 'pagamentoId', 'pagoEm', 'obs',
                  'categoria', 'responsavelId', 'responsavel', 'cliente', 'clienteContato', 'clienteDoc', 'formaPagamento', 'unidade', 'acao'],
  Pagamentos:    ['id', 'dataHora', 'colaboradorId', 'colaborador', 'valor', 'retiradaIds', 'forma',
                  'comprovanteId', 'status', 'conferidoEm', 'obs', 'categoria', 'responsavel', 'cliente'],
  Movimentos:    ['id', 'dataHora', 'itemId', 'item', 'tipo', 'qtd', 'saldo', 'obs']
};

/* ===================== INSTALAÇÃO ===================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abra o Apps Script pela planilha (Extensões > Apps Script).');
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  Object.keys(ABAS).forEach(function (n) {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, ABAS[n].length).setValues([ABAS[n]])
        .setFontWeight('bold').setBackground('#0a8f4e').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    } else {
      // Atualização: acrescenta no fim as colunas novas, sem mexer nos dados existentes
      const atual = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      ABAS[n].forEach(function (col, j) {
        if (atual[j] !== col) {
          if (atual.indexOf(col) >= 0) throw new Error('A aba ' + n + ' está com as colunas fora de ordem. Fale com o suporte antes de continuar.');
          if (atual[j] === undefined || atual[j] === '') {
            sh.getRange(1, j + 1).setValue(col).setFontWeight('bold').setBackground('#0a8f4e').setFontColor('#ffffff');
            atual[j] = col;
          } else throw new Error('A aba ' + n + ' tem uma coluna inesperada: ' + atual[j]);
        }
      });
    }
  });
  ['Página1', 'Planilha1', 'Sheet1'].forEach(function (x) {
    const s = ss.getSheetByName(x);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  // Preenche a unidade dos colaboradores antigos e cria a conta compartilhada de Ações Externas
  const shC = ss.getSheetByName('Colaboradores'), hC = ABAS.Colaboradores;
  const vC = shC.getDataRange().getValues();
  let temAcoes = false;
  for (let i = 1; i < vC.length; i++) {
    if (vC[i].join('') === '') continue;
    const un = vC[i][hC.indexOf('unidade')], setor = vC[i][hC.indexOf('setor')];
    if (un === UN.ACOES) temAcoes = true;
    if (!un) shC.getRange(i + 1, hC.indexOf('unidade') + 1).setValue(setor === 'Clínica' ? UN.AMOR : UN.CDT);
  }
  if (!temAcoes) {
    shC.appendRow(hC.map(function (k) {
      return ({ id: id_('C'), nome: NOME_CONTA_ACOES, setor: UN.ACOES, pinHash: '', ativo: 'SIM', criadoEm: agora_(), unidade: UN.ACOES })[k] || '';
    }));
  }
  pastaRaiz_();
  Logger.log('✅ Tudo pronto! Agora implante como App da Web.');
}

/* ===================== ROTEADOR ===================== */
function doGet() {
  return json_({ ok: true, msg: 'API do Mercadinho de TODOS ativa.' });
}

function doPost(e) {
  const ACOES = {
    catalogo: catalogo, login: login, retirar: retirar, minhaConta: minhaConta, pagar: pagar,
    adminDados: adminDados, salvarParceiro: salvarParceiro, salvarItem: salvarItem, movEstoque: movEstoque,
    salvarColaborador: salvarColaborador, resetarPin: resetarPin, cadastrar: cadastrar, definirPin: definirPin, conferirPagamento: conferirPagamento,
    baixaManual: baixaManual, cancelarRetirada: cancelarRetirada
  };
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, erro: 'Requisição inválida.' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const fn = ACOES[req.acao];
    if (!fn) throw new Error('Ação desconhecida.');
    return json_(Object.assign({ ok: true }, fn(req) || {}));
  } catch (err) {
    return json_({ ok: false, erro: err.message || String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ===================== PÚBLICO / COLABORADOR ===================== */
function catalogo() {
  const parc = ler_('Parceiros').filter(function (p) { return sim_(p.ativo); });
  const ids = parc.map(function (p) { return String(p.id); });
  return {
    parceiros: parc.map(function (p) { return { id: String(p.id), nome: p.nome, pagamento: p.pagamento }; }),
    itens: ler_('Itens').filter(function (i) { return sim_(i.ativo) && ids.indexOf(String(i.parceiroId)) >= 0; })
      .map(function (i) { return { id: String(i.id), parceiroId: String(i.parceiroId), nome: i.nome, preco: din_(i.preco), estoque: num_(i.estoque) }; }),
    colaboradores: ler_('Colaboradores').filter(function (c) { return sim_(c.ativo); })
      .map(function (c) { return { id: String(c.id), nome: c.nome, setor: c.setor, unidade: unidade_(c), primeiroAcesso: !c.pinHash }; })
      .sort(function (a, b) { return String(a.nome).localeCompare(String(b.nome)); })
  };
}

function login(req) {
  const c = acharColab_(req.colabId);
  if (!c.pinHash) {
    if (!req.criarPin) return { primeiroAcesso: true };
    if (!/^\d{4}$/.test(String(req.pin || ''))) throw new Error('O PIN precisa ter 4 números.');
    upd_('Colaboradores', c._row, { pinHash: hash_(req.pin) });
  } else if (c.pinHash !== hash_(req.pin)) {
    throw new Error('PIN incorreto.');
  }
  return { colab: { id: String(c.id), nome: c.nome, setor: c.setor, unidade: unidade_(c) } };
}

function cadastrar(req) {
  const nome = String(req.nome || '').trim().replace(/\s+/g, ' ');
  const un = req.unidade === UN.AMOR ? UN.AMOR : (req.unidade === UN.CDT ? UN.CDT : '');
  if (!un) throw new Error('Escolha a unidade.');
  if (nome.length < 3) throw new Error('Informe seu nome completo.');
  const setor = un === UN.AMOR ? 'Clínica' : String(req.setor || '');
  if (un === UN.CDT && (SETORES.indexOf(setor) < 0 || setor === 'Clínica')) throw new Error('Escolha o seu setor.');
  if (!/^\d{4}$/.test(String(req.pin || ''))) throw new Error('O PIN precisa ter 4 números.');
  const lista = ler_('Colaboradores');
  if (lista.some(function (x) { return String(x.nome).trim().toLowerCase() === nome.toLowerCase(); }))
    throw new Error('Já existe um cadastro com esse nome. Fale com o Thiago (Parcerias).');
  add_('Colaboradores', { id: id_('C'), nome: nome, setor: setor, pinHash: hash_(req.pin), ativo: APROVAR_CADASTROS ? 'PENDENTE' : 'SIM',
    criadoEm: agora_(), unidade: un, whatsapp: String(req.whatsapp || '').trim(), funcao: String(req.funcao || '').trim().slice(0, 80) });
  return { pendente: APROVAR_CADASTROS };
}

function retirar(req) {
  const c = colab_(req);
  const lista = (req.itens || []).filter(function (x) { return num_(x.qtd) > 0; });
  if (!lista.length) throw new Error('Escolha pelo menos um item.');
  const catKey = CATEGORIAS[req.categoria] ? req.categoria : 'COLABORADOR';
  const venda = catKey !== 'COLABORADOR';
  const unC = unidade_(c), contaAcoes = unC === UN.ACOES;
  if (contaAcoes && !venda) throw new Error('No acesso de Ações Externas use "Registrar venda".');
  if (unC === UN.AMOR && venda) throw new Error('Registro de vendas não está disponível para a sua unidade.');
  const vendedor = contaAcoes ? String(req.vendedor || '').trim() : c.nome;
  const acao = String(req.acao || '').trim().slice(0, 120);
  if (contaAcoes && !vendedor) throw new Error('Informe o seu nome (quem está vendendo).');
  if (contaAcoes && !acao) throw new Error('Informe a ação ou evento.');
  const obs = String(req.obs || '').trim().slice(0, 500);
  let cli = {};
  if (venda) {
    cli = { nome: String(req.cliente || '').trim(), contato: String(req.clienteContato || '').trim(), doc: String(req.clienteDoc || '').trim() };
    if (!cli.nome) throw new Error('Informe o nome do cliente.');
    if (FORMAS.indexOf(req.formaPagamento) < 0) throw new Error('Escolha a forma de pagamento.');
    if (req.formaPagamento !== 'Dinheiro' && !req.comprovante) throw new Error('Anexe o comprovante de pagamento.');
  } else if (!req.foto) throw new Error('A foto dos itens é obrigatória.');
  const modo = venda ? 'pago' : (req.modo === 'pago' ? 'pago' : 'fiado');
  if (!venda && modo === 'fiado' && unC !== UN.CDT)
    throw new Error('Para a sua unidade o pagamento é no ato: escolha "Já paguei" e anexe o comprovante.');
  if (!venda && modo === 'pago' && !req.comprovante) throw new Error('Anexe o comprovante de pagamento.');

  const itens = ler_('Itens'), parceiros = ler_('Parceiros');
  const linhas = lista.map(function (x) {
    const it = itens.filter(function (i) { return String(i.id) === String(x.itemId); })[0];
    if (!it || !sim_(it.ativo)) throw new Error('Um dos itens não está mais disponível.');
    const q = Math.floor(num_(x.qtd));
    if (num_(it.estoque) < q) throw new Error('Estoque insuficiente de "' + it.nome + '" (restam ' + num_(it.estoque) + ').');
    return { it: it, q: q };
  });
  const pids = linhas.map(function (l) { return String(l.it.parceiroId); })
    .filter(function (v, i, a) { return a.indexOf(v) === i; });
  if (pids.length > 1) throw new Error('Faça uma retirada por parceiro.');
  const parc = parceiros.filter(function (p) { return String(p.id) === pids[0]; })[0] || {};

  const agora = agora_(), pedidoId = id_('PD');
  const total = din_(linhas.reduce(function (s, l) { return s + l.q * num_(l.it.preco); }, 0));
  const quem = venda ? cli.nome : c.nome;
  const fotoId = req.foto ? salvarArquivo_(req.foto, nomeArq_(agora, quem, 'retirada'), 'Fotos das retiradas') : '';
  let compId = '', pagId = '';
  if (modo === 'pago') {
    if (req.comprovante) compId = salvarArquivo_(req.comprovante, nomeArq_(agora, quem, 'comprovante'), 'Comprovantes');
    pagId = id_('G');
  }
  const ids = [];
  linhas.forEach(function (l) {
    const novo = num_(l.it.estoque) - l.q;
    upd_('Itens', l.it._row, { estoque: novo });
    add_('Movimentos', { id: id_('M'), dataHora: agora, itemId: l.it.id, item: l.it.nome, tipo: 'RETIRADA', qtd: -l.q, saldo: novo, obs: (venda ? CATEGORIAS[catKey] + ': ' + cli.nome + ' (resp. ' + vendedor + ')' : c.nome) + ' · ' + pedidoId });
    const rid = id_('R');
    ids.push(rid);
    add_('Retiradas', {
      id: rid, pedidoId: pedidoId, dataHora: agora,
      colaboradorId: venda ? '' : c.id, colaborador: venda ? cli.nome : c.nome, setor: venda ? CATEGORIAS[catKey] : c.setor,
      parceiroId: parc.id, parceiro: parc.nome, itemId: l.it.id, item: l.it.nome, qtd: l.q,
      precoUnit: din_(l.it.preco), total: din_(l.q * num_(l.it.preco)), fotoId: fotoId,
      modo: venda ? 'VENDA' : (modo === 'pago' ? 'PAGO NA HORA' : 'FIADO'), status: modo === 'pago' ? 'EM ANÁLISE' : 'PENDENTE',
      pagamentoId: pagId, pagoEm: '', obs: obs,
      categoria: CATEGORIAS[catKey], responsavelId: c.id, responsavel: vendedor,
      unidade: !venda ? unC : (catKey === 'AVULSA' ? UN.ACOES : UN.CDT), acao: acao,
      cliente: venda ? cli.nome : '', clienteContato: venda ? cli.contato : '', clienteDoc: venda ? cli.doc : '',
      formaPagamento: venda ? req.formaPagamento : (modo === 'pago' ? 'Comprovante' : '')
    });
  });
  if (pagId) {
    add_('Pagamentos', { id: pagId, dataHora: agora, colaboradorId: venda ? '' : c.id, colaborador: venda ? cli.nome : c.nome, valor: total, retiradaIds: ids.join(','),
      forma: venda ? req.formaPagamento : 'Comprovante na retirada', comprovanteId: compId, status: 'EM ANÁLISE', conferidoEm: '', obs: obs,
      categoria: CATEGORIAS[catKey], responsavel: vendedor, cliente: venda ? cli.nome : '' });
  }
  return {
    pedidoId: pedidoId, total: total, parceiro: parc.nome, dataHora: agora, categoria: CATEGORIAS[catKey], cliente: venda ? cli.nome : '', vendedor: vendedor, acao: acao,
    itens: linhas.map(function (l) { return { nome: l.it.nome, qtd: l.q, total: din_(l.q * num_(l.it.preco)) }; })
  };
}

function minhaConta(req) {
  const c = colab_(req);
  const meu = function (x) { return String(x.colaboradorId) === String(c.id); };
  return {
    retiradas: ler_('Retiradas').filter(meu).map(limpa_).reverse(),
    vendas: ler_('Retiradas').filter(function (x) { return String(x.responsavelId) === String(c.id) && x.categoria && x.categoria !== CATEGORIAS.COLABORADOR; }).map(limpa_).reverse(),
    pagamentos: ler_('Pagamentos').filter(meu).map(limpa_).reverse(),
    parceiros: ler_('Parceiros').map(function (p) { return { id: String(p.id), nome: p.nome, pagamento: p.pagamento }; })
  };
}

function pagar(req) {
  const c = colab_(req);
  const ids = (req.retiradaIds || []).map(String);
  if (!ids.length) throw new Error('Selecione o que você está pagando.');
  if (!req.comprovante) throw new Error('Anexe o comprovante.');
  const rs = ler_('Retiradas').filter(function (r) { return ids.indexOf(String(r.id)) >= 0; });
  if (rs.length !== ids.length) throw new Error('Registro não encontrado. Atualize a página.');
  rs.forEach(function (r) {
    if (String(r.colaboradorId) !== String(c.id)) throw new Error('Registro de outro colaborador.');
    if (r.status !== 'PENDENTE') throw new Error('"' + r.item + '" não está mais em aberto. Atualize a página.');
  });
  const agora = agora_(), pid = id_('G');
  const valor = din_(rs.reduce(function (s, r) { return s + num_(r.total); }, 0));
  const comp = salvarArquivo_(req.comprovante, nomeArq_(agora, c.nome, 'comprovante'), 'Comprovantes');
  add_('Pagamentos', { id: pid, dataHora: agora, colaboradorId: c.id, colaborador: c.nome, valor: valor, retiradaIds: ids.join(','),
    forma: 'Comprovante enviado', comprovanteId: comp, status: 'EM ANÁLISE', conferidoEm: '', obs: String(req.obs || '').trim().slice(0, 500),
    categoria: CATEGORIAS.COLABORADOR, responsavel: c.nome, cliente: '' });
  rs.forEach(function (r) { upd_('Retiradas', r._row, { status: 'EM ANÁLISE', pagamentoId: pid }); });
  return { pagamentoId: pid, valor: valor };
}

/* ===================== GESTOR ===================== */
function adminDados(req) {
  admin_(req);
  return {
    parceiros: ler_('Parceiros').map(limpa_),
    itens: ler_('Itens').map(limpa_),
    colaboradores: ler_('Colaboradores').map(function (c) {
      return { id: String(c.id), nome: c.nome, setor: c.setor, ativo: c.ativo, criadoEm: c.criadoEm, temPin: !!c.pinHash,
        unidade: unidade_(c), whatsapp: c.whatsapp || '', funcao: c.funcao || '' };
    }),
    retiradas: ler_('Retiradas').map(limpa_),
    pagamentos: ler_('Pagamentos').map(limpa_),
    movimentos: ler_('Movimentos').slice(-200).map(limpa_)
  };
}

function salvarParceiro(req) {
  admin_(req);
  const p = req.parceiro || {};
  const nome = String(p.nome || '').trim();
  if (!nome) throw new Error('Informe o nome do parceiro.');
  const lista = ler_('Parceiros');
  if (lista.some(function (x) { return String(x.id) !== String(p.id || '') && String(x.nome).trim().toLowerCase() === nome.toLowerCase(); }))
    throw new Error('Já existe um parceiro com esse nome.');
  const dados = { nome: nome, contato: p.contato || '', whatsapp: p.whatsapp || '', pagamento: p.pagamento || '', ativo: p.ativo === 'NÃO' ? 'NÃO' : 'SIM' };
  if (p.id) {
    const a = lista.filter(function (x) { return String(x.id) === String(p.id); })[0];
    if (!a) throw new Error('Parceiro não encontrado.');
    upd_('Parceiros', a._row, dados);
  } else {
    add_('Parceiros', Object.assign({ id: id_('P'), criadoEm: agora_() }, dados));
  }
  return {};
}

function salvarItem(req) {
  admin_(req);
  const i = req.item || {};
  const nome = String(i.nome || '').trim();
  if (!nome) throw new Error('Informe o nome do item.');
  if (!(num_(i.preco) > 0)) throw new Error('Informe um preço maior que zero.');
  if (!ler_('Parceiros').some(function (p) { return String(p.id) === String(i.parceiroId); })) throw new Error('Escolha o parceiro.');
  const lista = ler_('Itens');
  if (lista.some(function (x) { return String(x.id) !== String(i.id || '') && String(x.parceiroId) === String(i.parceiroId) && String(x.nome).trim().toLowerCase() === nome.toLowerCase(); }))
    throw new Error('Esse parceiro já tem um item com esse nome.');
  const dados = { parceiroId: String(i.parceiroId), nome: nome, preco: din_(i.preco), estoqueMinimo: Math.max(0, Math.floor(num_(i.estoqueMinimo))), ativo: i.ativo === 'NÃO' ? 'NÃO' : 'SIM' };
  if (i.id) {
    const a = lista.filter(function (x) { return String(x.id) === String(i.id); })[0];
    if (!a) throw new Error('Item não encontrado.');
    upd_('Itens', a._row, dados);
  } else {
    const est = Math.max(0, Math.floor(num_(i.estoqueInicial)));
    const nid = id_('I'), agora = agora_();
    add_('Itens', Object.assign({ id: nid, estoque: est, criadoEm: agora }, dados));
    if (est > 0) add_('Movimentos', { id: id_('M'), dataHora: agora, itemId: nid, item: nome, tipo: 'ENTRADA', qtd: est, saldo: est, obs: 'Estoque inicial' });
  }
  return {};
}

function movEstoque(req) {
  admin_(req);
  const it = ler_('Itens').filter(function (x) { return String(x.id) === String(req.itemId); })[0];
  if (!it) throw new Error('Item não encontrado.');
  const est = num_(it.estoque), q = Math.floor(num_(req.qtd));
  let novo, delta;
  if (req.tipo === 'ENTRADA') {
    if (!(q > 0)) throw new Error('Informe uma quantidade maior que zero.');
    delta = q; novo = est + q;
  } else if (req.tipo === 'AJUSTE') {
    if (q < 0 || String(req.qtd).trim() === '') throw new Error('Informe a contagem real (0 ou mais).');
    if (!String(req.obs || '').trim()) throw new Error('Informe o motivo do ajuste.');
    novo = q; delta = novo - est;
  } else throw new Error('Tipo de movimento inválido.');
  upd_('Itens', it._row, { estoque: novo });
  add_('Movimentos', { id: id_('M'), dataHora: agora_(), itemId: it.id, item: it.nome, tipo: req.tipo, qtd: delta, saldo: novo, obs: req.obs || '' });
  return { saldo: novo };
}

function salvarColaborador(req) {
  admin_(req);
  const c = req.colab || {};
  const nome = String(c.nome || '').trim();
  if (!nome) throw new Error('Informe o nome do colaborador.');
  const un = [UN.CDT, UN.AMOR, UN.ACOES].indexOf(c.unidade) >= 0 ? c.unidade : UN.CDT;
  if (un === UN.AMOR) c.setor = 'Clínica';
  if (un === UN.ACOES) c.setor = UN.ACOES;
  else if (SETORES.indexOf(c.setor) < 0) throw new Error('Escolha o setor.');
  const lista = ler_('Colaboradores');
  if (lista.some(function (x) { return String(x.id) !== String(c.id || '') && String(x.nome).trim().toLowerCase() === nome.toLowerCase(); }))
    throw new Error('Já existe um colaborador com esse nome.');
  const dados = { nome: nome, setor: c.setor, ativo: c.ativo === 'NÃO' ? 'NÃO' : 'SIM', unidade: un };
  if (c.whatsapp !== undefined) dados.whatsapp = String(c.whatsapp || '').trim();
  if (c.funcao !== undefined) dados.funcao = String(c.funcao || '').trim().slice(0, 80);
  if (c.id) {
    const a = lista.filter(function (x) { return String(x.id) === String(c.id); })[0];
    if (!a) throw new Error('Colaborador não encontrado.');
    upd_('Colaboradores', a._row, dados);
  } else {
    add_('Colaboradores', Object.assign({ id: id_('C'), pinHash: '', criadoEm: agora_() }, dados));
  }
  return {};
}

function definirPin(req) {
  admin_(req);
  const c = acharColab_(req.colabId, true);
  if (!/^\d{4}$/.test(String(req.pin || ''))) throw new Error('A senha precisa ter 4 números.');
  upd_('Colaboradores', c._row, { pinHash: hash_(req.pin) });
  return {};
}

function resetarPin(req) {
  admin_(req);
  const c = acharColab_(req.colabId, true);
  upd_('Colaboradores', c._row, { pinHash: '' });
  return {};
}

function conferirPagamento(req) {
  admin_(req);
  const p = ler_('Pagamentos').filter(function (x) { return String(x.id) === String(req.pagamentoId); })[0];
  if (!p) throw new Error('Pagamento não encontrado.');
  if (p.status !== 'EM ANÁLISE') throw new Error('Este pagamento já foi conferido.');
  const ids = String(p.retiradaIds).split(',').filter(String);
  const rs = ler_('Retiradas').filter(function (r) { return ids.indexOf(String(r.id)) >= 0; });
  const agora = agora_();
  if (req.decisao === 'CONFIRMAR') {
    upd_('Pagamentos', p._row, { status: 'CONFIRMADO', conferidoEm: agora, obs: req.obs || '' });
    rs.forEach(function (r) { upd_('Retiradas', r._row, { status: 'PAGO', pagoEm: agora }); });
  } else if (req.decisao === 'RECUSAR') {
    if (!String(req.obs || '').trim()) throw new Error('Informe o motivo da recusa.');
    upd_('Pagamentos', p._row, { status: 'RECUSADO', conferidoEm: agora, obs: req.obs });
    rs.forEach(function (r) { upd_('Retiradas', r._row, { status: 'PENDENTE', pagamentoId: '' }); });
  } else throw new Error('Decisão inválida.');
  return {};
}

function baixaManual(req) {
  admin_(req);
  const ids = (req.retiradaIds || []).map(String);
  if (!ids.length) throw new Error('Selecione as retiradas.');
  if (!req.forma) throw new Error('Informe a forma de pagamento.');
  const rs = ler_('Retiradas').filter(function (r) { return ids.indexOf(String(r.id)) >= 0; });
  rs.forEach(function (r) { if (r.status !== 'PENDENTE') throw new Error('"' + r.item + '" de ' + r.colaborador + ' não está pendente.'); });
  const agora = agora_(), grupos = {};
  rs.forEach(function (r) { const k = r.colaboradorId ? 'C:' + r.colaboradorId : 'P:' + r.pedidoId; (grupos[k] = grupos[k] || []).push(r); });
  Object.keys(grupos).forEach(function (cid) {
    const g = grupos[cid], pid = id_('G');
    add_('Pagamentos', { id: pid, dataHora: agora, colaboradorId: g[0].colaboradorId || '', colaborador: g[0].colaborador,
      valor: din_(g.reduce(function (s, r) { return s + num_(r.total); }, 0)), retiradaIds: g.map(function (r) { return r.id; }).join(','),
      forma: 'Baixa manual: ' + req.forma, comprovanteId: '', status: 'CONFIRMADO', conferidoEm: agora, obs: req.obs || '',
      categoria: g[0].categoria || CATEGORIAS.COLABORADOR, responsavel: g[0].responsavel || '', cliente: g[0].cliente || '' });
    g.forEach(function (r) { upd_('Retiradas', r._row, { status: 'PAGO', pagoEm: agora, pagamentoId: pid }); });
  });
  return {};
}

function cancelarRetirada(req) {
  admin_(req);
  const r = ler_('Retiradas').filter(function (x) { return String(x.id) === String(req.retiradaId); })[0];
  if (!r) throw new Error('Retirada não encontrada.');
  if (r.status !== 'PENDENTE') throw new Error('Só é possível cancelar retiradas em aberto.');
  if (!String(req.obs || '').trim()) throw new Error('Informe o motivo do cancelamento.');
  upd_('Retiradas', r._row, { status: 'CANCELADA', obs: (r.obs ? r.obs + ' | ' : '') + 'Cancelada: ' + req.obs });
  const it = ler_('Itens').filter(function (x) { return String(x.id) === String(r.itemId); })[0];
  if (it) {
    const novo = num_(it.estoque) + num_(r.qtd);
    upd_('Itens', it._row, { estoque: novo });
    add_('Movimentos', { id: id_('M'), dataHora: agora_(), itemId: it.id, item: it.nome, tipo: 'ESTORNO', qtd: num_(r.qtd), saldo: novo, obs: 'Cancelamento ' + r.id + ' · ' + req.obs });
  }
  return {};
}

/* ===================== AUXILIARES ===================== */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Sistema não instalado: rode a função setup.');
  return SpreadsheetApp.openById(id);
}
function sh_(n) { return ss_().getSheetByName(n); }
function ler_(n) {
  const v = sh_(n).getDataRange().getValues();
  const h = v.shift() || [];
  return v.filter(function (r) { return r.join('') !== ''; }).map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) {
      let x = r[j];
      if (x instanceof Date) x = Utilities.formatDate(x, TZ, 'yyyy-MM-dd HH:mm');
      o[k] = x;
    });
    return o;
  });
}
function seguro_(v) { return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v; }
function add_(n, o) { sh_(n).appendRow(ABAS[n].map(function (k) { return o[k] === undefined ? '' : seguro_(o[k]); })); }
function upd_(n, row, patch) {
  const sh = sh_(n), h = ABAS[n];
  Object.keys(patch).forEach(function (k) {
    const c = h.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(seguro_(patch[k]));
  });
}
function limpa_(o) { const c = Object.assign({}, o); delete c._row; Object.keys(c).forEach(function (k) { if (/id$|Id$|Ids$/.test(k)) c[k] = String(c[k]); }); return c; }
function agora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function id_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(); }
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function din_(v) { return Math.round(num_(v) * 100) / 100; }
function sim_(v) { return String(v).toUpperCase() === 'SIM'; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hash_(s) {
  return 'h' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s) + '|mercadinho-cdt')
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function admin_(req) { if (String(req.senha || '') !== SENHA_ADMIN) throw new Error('Senha do gestor incorreta.'); }
function acharColab_(id, mesmoInativo) {
  const c = ler_('Colaboradores').filter(function (x) { return String(x.id) === String(id); })[0];
  if (c && !mesmoInativo && String(c.ativo).toUpperCase() === 'PENDENTE') throw new Error('Seu cadastro está aguardando aprovação do Thiago (Parcerias).');
  if (!c || (!mesmoInativo && !sim_(c.ativo))) throw new Error('Colaborador não encontrado ou inativo.');
  return c;
}
function unidade_(c) {
  if (c.unidade) return c.unidade;
  return c.setor === 'Clínica' ? UN.AMOR : UN.CDT;
}
function colab_(req) {
  const c = acharColab_(req.colabId);
  if (!c.pinHash || c.pinHash !== hash_(req.pin)) throw new Error('PIN incorreto. Entre novamente.');
  return c;
}
function pastaRaiz_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PASTA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Mercadinho de TODOS - Fotos e comprovantes');
  props.setProperty('PASTA_ID', f.getId());
  return f;
}
function subpasta_(nome) {
  const r = pastaRaiz_(), it = r.getFoldersByName(nome);
  return it.hasNext() ? it.next() : r.createFolder(nome);
}
function nomeArq_(agora, pessoa, tipo) {
  return (agora.replace(/[: ]/g, '-') + '_' + pessoa + '_' + tipo).replace(/[^\wÀ-ú-]+/g, '_');
}
function salvarArquivo_(dataUrl, base, sub) {
  const m = String(dataUrl || '').match(/^data:([\w\/+.-]+);base64,(.+)$/);
  if (!m) throw new Error('Arquivo inválido. Tente anexar de novo.');
  const tipos = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'application/pdf': 'pdf' };
  if (!tipos[m[1]]) throw new Error('Envie uma foto ou um PDF.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 6 * 1024 * 1024) throw new Error('Arquivo muito grande (máx. 6 MB).');
  const f = subpasta_(sub).createFile(Utilities.newBlob(bytes, m[1], base + '.' + tipos[m[1]]));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return f.getId();
}

/* ===================== CARGA INICIAL (opcional) =====================
 * Preencha as listas e rode a função  cargaInicial  UMA vez.
 * Cada sabor vira um item próprio, para o estoque ficar certo por sabor.
 * Colaboradores entram sem PIN: no 1º acesso cada um cria o seu.
 * =================================================================== */
const PRECO_BROWNIE = 13.90; // ← coloque aqui o preço do brownie da Janna Gourmet (ex.: 8.50)

const CARGA = {
  parceiros: [
    { nome: 'Supremo Geladinho Gourmet', contato: '', whatsapp: '', pagamento: 'Pague pelo app do Cartão de TODOS (cartão de crédito, Pix ou saldo de cashback) e anexe o comprovante aqui.' },
    { nome: 'Janna Gourmet', contato: '', whatsapp: '', pagamento: 'Pague pelo app do Cartão de TODOS (cartão de crédito, Pix ou saldo de cashback) e anexe o comprovante aqui.' }
  ],
  itens: [
    { parceiro: 'Supremo Geladinho Gourmet', nome: 'Supremo – Ninho com Nutella',   preco: 12, estoque: 10, minimo: 3 },
    { parceiro: 'Supremo Geladinho Gourmet', nome: 'Supremo – Morango com Nutella', preco: 12, estoque: 10, minimo: 3 },
    { parceiro: 'Supremo Geladinho Gourmet', nome: 'Gourmet – Maracujá Trufado',    preco: 12, estoque: 10, minimo: 3 },
    { parceiro: 'Supremo Geladinho Gourmet', nome: 'Tradicional – Abacaxi Cremoso', preco: 9,  estoque: 10, minimo: 3 },
    { parceiro: 'Supremo Geladinho Gourmet', nome: 'Tradicional – Brigadeiro',      preco: 9,  estoque: 10, minimo: 3 },
    { parceiro: 'Janna Gourmet', nome: 'Brownie – Ninho com Nutella', preco: PRECO_BROWNIE, estoque: 15, minimo: 3 },
    { parceiro: 'Janna Gourmet', nome: 'Brownie – Doce de Leite',     preco: PRECO_BROWNIE, estoque: 15, minimo: 3 }
  ],
  colaboradores: [
    { nome: 'Janderson', setor: 'Vendas' },
    { nome: 'Joyce', setor: 'Vendas' },
    { nome: 'Jéssica Posca', setor: 'Vendas' },
    { nome: 'Daniele Fernandes', setor: 'Vendas' },
    { nome: 'Poliana', setor: 'Vendas' },
    { nome: 'Leandro', setor: 'Vendas' },
    { nome: 'Denilson', setor: 'Retenção' },
    { nome: 'Deidiane', setor: 'Administrativo' },
    { nome: 'Luiz', setor: 'Administrativo' },
    { nome: 'Jéssica Natiele', setor: 'Adimplência' },
    { nome: 'Ana Beatriz', setor: 'Adimplência' },
    { nome: 'Janaína', setor: 'Adimplência' },
    { nome: 'Thiago', setor: 'Parcerias' },
    { nome: 'Rodolfo', setor: 'Gerência' }
  ]
};

function cargaInicial() {
  const req = { senha: SENHA_ADMIN };
  const jaP = ler_('Parceiros').map(function (p) { return String(p.nome).toLowerCase(); });
  CARGA.parceiros.forEach(function (p) {
    if (jaP.indexOf(p.nome.toLowerCase()) < 0) salvarParceiro(Object.assign({}, req, { parceiro: p }));
  });
  const parc = ler_('Parceiros');
  const jaI = ler_('Itens').map(function (i) { return i.parceiroId + '|' + String(i.nome).toLowerCase(); });
  CARGA.itens.forEach(function (i) {
    const p = parc.filter(function (x) { return String(x.nome).toLowerCase() === i.parceiro.toLowerCase(); })[0];
    if (!p) throw new Error('Parceiro não encontrado na carga: ' + i.parceiro);
    if (jaI.indexOf(p.id + '|' + i.nome.toLowerCase()) >= 0) return;
    if (!(num_(i.preco) > 0)) { Logger.log('⚠️ Pulei "' + i.nome + '": falta o preço. Preencha e rode de novo.'); return; }
    salvarItem(Object.assign({}, req, { item: { parceiroId: p.id, nome: i.nome, preco: i.preco, estoqueInicial: i.estoque, estoqueMinimo: i.minimo || 0 } }));
  });
  const jaC = ler_('Colaboradores').map(function (c) { return String(c.nome).toLowerCase(); });
  CARGA.colaboradores.forEach(function (c) {
    if (jaC.indexOf(c.nome.toLowerCase()) < 0) salvarColaborador(Object.assign({}, req, { colab: c }));
  });
  Logger.log('✅ Carga inicial concluída.');
}
