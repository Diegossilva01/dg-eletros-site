const API_URL="https://script.google.com/macros/s/AKfycbwbh1LHJyEVyV41q_7u5U3crbZ-zGIZKeu4gWuaOpTNKykHeAeUpDnIkS62ZEDzt4EG/exec";
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const money=v=>Number(v||0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
const state={token:localStorage.getItem('dge_token')||'',user:null,produtos:[],vendas:[],usuarios:[],comissoes:[],receber:[],horasExtras:[]};
const receberExpandidos=new Set();

function protegerSessaoEInterface(){
  document.querySelectorAll('button:not([type])').forEach(botao=>{
    if(!botao.closest('form')) botao.type='button';
  });
  document.addEventListener('click',e=>{
    const nav=e.target.closest('.nav[data-page]');
    if(nav)localStorage.setItem('dge_pagina_ativa',nav.dataset.page||'dashboard');
  });
  window.addEventListener('beforeunload',()=>{
    if(state.token)localStorage.setItem('dge_token',state.token);
  });
}
function restaurarPaginaAtiva(){
  const pagina=localStorage.getItem('dge_pagina_ativa');
  if(!pagina||isRevendedor())return;
  const botao=document.querySelector(`.nav[data-page="${pagina}"]`);
  if(botao&&!botao.classList.contains('hidden'))setTimeout(()=>botao.click(),0);
}

function aplicarTema(tema){const escuro=tema==='dark';document.body.classList.toggle('dark-theme',escuro);localStorage.setItem('dge_tema',escuro?'dark':'light');const b=$('#themeToggle');if(b)b.innerHTML=escuro?'<i class="fa-solid fa-sun"></i><span>Tema claro</span>':'<i class="fa-solid fa-moon"></i><span>Tema escuro</span>';}
function iniciarTema(){aplicarTema(localStorage.getItem('dge_tema')||'light');const b=$('#themeToggle');if(b)b.onclick=()=>aplicarTema(document.body.classList.contains('dark-theme')?'light':'dark');}

const API_TIMEOUT_MS=30000;
const API_TENTATIVAS=3;
const esperar=ms=>new Promise(resolve=>setTimeout(resolve,ms));

function mensagemErroRede(err){
  const texto=String(err?.message||err||'').trim();
  if(err?.name==='AbortError')return 'O servidor demorou para responder. Tente novamente.';
  if(/load failed|failed to fetch|networkerror|network request failed|string did not match/i.test(texto))return 'Falha temporária de conexão com o servidor.';
  return texto||'Falha ao acessar o servidor.';
}

async function api(acao,dados={},opcoes={}){
  let url;
  try{url=new URL(String(API_URL).trim()).toString()}catch(_){throw new Error('A URL da API está inválida. Verifique o endereço do Google Apps Script.')}
  let ultimoErro;
  const tentativas=Number(opcoes.tentativas||API_TENTATIVAS);
  for(let tentativa=1;tentativa<=tentativas;tentativa++){
    const controller=new AbortController();
    const timeout=setTimeout(()=>controller.abort(),API_TIMEOUT_MS);
    try{
      const r=await fetch(url,{
        method:'POST',
        headers:{'Content-Type':'text/plain;charset=utf-8'},
        body:JSON.stringify({acao,token:state.token,...dados}),
        cache:'no-store',
        redirect:'follow',
        signal:controller.signal
      });
      const texto=await r.text();
      if(!r.ok)throw new Error(`Servidor indisponível (${r.status}).`);
      let j;
      try{j=JSON.parse(texto)}catch(_){throw new Error('O servidor retornou uma resposta inválida.')}
      if(!j.ok){
        const erro=new Error(j.erro||'Erro na operação');
        erro.naoRepetir=true;
        throw erro;
      }
      return j;
    }catch(err){
      ultimoErro=err;
      if(err?.naoRepetir||tentativa>=tentativas)break;
      await esperar(700*tentativa);
    }finally{clearTimeout(timeout)}
  }
  throw new Error(mensagemErroRede(ultimoErro));
}
function hoje(){return new Date().toISOString().slice(0,10)}
function parseMoney(v){return Number(String(v||'').replace(/R\$/g,'').replace(/\./g,'').replace(',','.').replace(/[^0-9.-]/g,''))||0}
function formObj(f){return Object.fromEntries(new FormData(f).entries())}
function isAdmin(){return state.user?.perfil==='Administrador'}
function isVendedor(){return ['Vendedora','Vendedor'].includes(state.user?.perfil)}
function podeGerenciarProdutos(){return isAdmin()||isVendedor()}
function isRevendedor(){return state.user?.perfil==='Revendedor'}
function somenteDigitos(v){return String(v||'').replace(/\D/g,'')}
function formatarCPFouCNPJ(v){
  const d=somenteDigitos(v).slice(0,14);
  if(d.length<=11)return d.replace(/(\d{3})(\d)/,'$1.$2').replace(/(\d{3})(\d)/,'$1.$2').replace(/(\d{3})(\d{1,2})$/,'$1-$2');
  return d.replace(/(\d{2})(\d)/,'$1.$2').replace(/(\d{3})(\d)/,'$1.$2').replace(/(\d{3})(\d)/,'$1/$2').replace(/(\d{4})(\d{1,2})$/,'$1-$2');
}
function formatarTelefone(v){
  const d=somenteDigitos(v).slice(0,11);
  if(d.length<=10)return d.replace(/(\d{2})(\d)/,'($1) $2').replace(/(\d{4})(\d{1,4})$/,'$1-$2');
  return d.replace(/(\d{2})(\d)/,'($1) $2').replace(/(\d{5})(\d{1,4})$/,'$1-$2');
}
function validarCPF(cpf){
  cpf=somenteDigitos(cpf);if(cpf.length!==11||/^(\d)\1{10}$/.test(cpf))return false;
  let soma=0;for(let i=0;i<9;i++)soma+=Number(cpf[i])*(10-i);let d=(soma*10)%11;if(d===10)d=0;if(d!==Number(cpf[9]))return false;
  soma=0;for(let i=0;i<10;i++)soma+=Number(cpf[i])*(11-i);d=(soma*10)%11;if(d===10)d=0;return d===Number(cpf[10]);
}
function validarCNPJ(cnpj){
  cnpj=somenteDigitos(cnpj);if(cnpj.length!==14||/^(\d)\1{13}$/.test(cnpj))return false;
  const calc=p=>{let soma=0;for(let i=0;i<p.length;i++)soma+=Number(cnpj[i])*p[i];const r=soma%11;return r<2?0:11-r};
  const d1=calc([5,4,3,2,9,8,7,6,5,4,3,2]);if(d1!==Number(cnpj[12]))return false;
  const d2=calc([6,5,4,3,2,9,8,7,6,5,4,3,2]);return d2===Number(cnpj[13]);
}
function validarDocumento(v){const d=somenteDigitos(v);return !d||d.length===11?(!d||validarCPF(d)):d.length===14&&validarCNPJ(d)}
function validarTelefone(v){const d=somenteDigitos(v);return !d||((d.length===10||d.length===11)&&d[0]!=='0'&&d[1]!=='0')}
function alternarTipoVenda(){
  const avulsa=$('#tipoVenda')?.value==='avulsa';
  $('#campoProdutoEstoque')?.classList.toggle('hidden',avulsa);
  $('#camposVendaAvulsa')?.classList.toggle('hidden',!avulsa);
  const sel=$('#produtoVenda');if(sel)sel.required=!avulsa;
  ['produtoManual','marcaManual','modeloManual','categoriaManual'].forEach((n,i)=>{const el=$(`#vendaForm [name=${n}]`);if(el)el.required=avulsa&&i===0});
}
const tipoVendaEl=$('#tipoVenda');if(tipoVendaEl){tipoVendaEl.addEventListener('change',alternarTipoVenda);alternarTipoVenda()}
const documentoEl=$('#documentoVenda');if(documentoEl)documentoEl.addEventListener('input',e=>e.target.value=formatarCPFouCNPJ(e.target.value));
const telefoneEl=$('#telefoneVenda');if(telefoneEl)telefoneEl.addEventListener('input',e=>e.target.value=formatarTelefone(e.target.value));
const produtoVendaEl=$('#produtoVenda');if(produtoVendaEl)produtoVendaEl.addEventListener('change',e=>{const op=e.target.selectedOptions[0];const linha=$('#vendaForm [name=linhaProduto]');if(linha)linha.value=op?.dataset.linha||'';if(op?.dataset.preco){const campo=$('#vendaForm [name=valorVenda]');if(campo&&!campo.value)campo.value=Number(op.dataset.preco).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}});


function atualizarSaldoVenda(){const total=parseMoney($('#vendaForm [name=valorVenda]')?.value),entrada=parseMoney($('#entradaVenda')?.value);const saldo=Math.max(0,total-entrada);if($('#saldoVenda'))$('#saldoVenda').value=money(saldo);return saldo}
function alternarPagamento(){const prazo=$('#pagamentoVenda')?.value==='No prazo';$('#prazoBox')?.classList.toggle('hidden',!prazo);const entrada=$('#entradaVenda');if(entrada)entrada.required=prazo;atualizarSaldoVenda()}
const pagamentoVenda=$('#pagamentoVenda');if(pagamentoVenda)pagamentoVenda.addEventListener('change',alternarPagamento);
$('#vendaForm [name=valorVenda]')?.addEventListener('input',atualizarSaldoVenda);$('#entradaVenda')?.addEventListener('input',atualizarSaldoVenda);
async function carregarReceber(){try{const j=await api('listarContasReceber');state.receber=j.contas||[];$('#totalReceber').textContent=money(j.totalReceber);$('#clientesPendentes').textContent=j.clientesPendentes||0;renderReceber()}catch(e){console.error(e)}}
function renderReceber(){const busca=normalizarTexto($('#buscaReceber')?.value);const lista=(state.receber||[]).filter(c=>!busca||normalizarTexto([c.cliente,c.telefone,c.produto].join(' ')).includes(busca));$('#receberBody').innerHTML=lista.map(c=>{const dias=Number(c.diasSemPagamento||0),faixa=dias<=7?'recent':dias<=30?'attention':'late',textoDias=dias===0?'Hoje':`Há ${dias} dia${dias===1?'':'s'}`;return `<tr class="payment-row-${faixa}"><td>${esc(c.dataVenda||'')}</td><td><strong>${esc(c.cliente)}</strong><br><small>${esc(c.telefone||'')}</small></td><td>${esc(c.produto)}</td><td>${money(c.valorVenda)}</td><td><div><strong>${money(c.totalPago||0)}</strong></div><small>${Number(c.quantidadePagamentos||0)} pagamento${Number(c.quantidadePagamentos||0)===1?'':'s'}</small></td><td><strong class="saldo-pendente">${money(c.saldo)}</strong></td><td><div class="last-payment-card ${faixa}"><span>Último pagamento</span><strong>${esc(c.ultimoPagamento||'Nenhum')}</strong><small>${esc(textoDias)}</small></div><button type="button" class="table-action action-paid js-receber-pagamento" data-venda="${esc(c.vendaId)}" data-saldo="${Number(c.saldo||0)}"><i class="fa-solid fa-minus"></i> Abater pagamento</button><button type="button" class="table-action action-receipt js-recibo-receber" data-venda="${esc(c.vendaId)}"><i class="fa-solid fa-file-invoice"></i> Recibo</button><details class="payment-history"><summary>Histórico (${Number(c.quantidadePagamentos||0)})</summary><div>${Number(c.entrada||0)>0?`<p><strong>${money(c.entrada)}</strong> — ${esc(c.dataVenda||'')} • Entrada da venda</p>`:''}${(c.recebimentos||[]).length?[...(c.recebimentos||[])].reverse().map(r=>`<p><strong>${money(r.valor)}</strong> — ${esc(r.data||'')} ${r.observacao?'• '+esc(r.observacao):''}${r.registradoPor?' • '+esc(r.registradoPor):''}</p>`).join(''):Number(c.entrada||0)>0?'':'<p>Nenhum pagamento registrado.</p>'}</div></details></td></tr>`}).join('')||'<tr><td colspan="7">Nenhum saldo pendente.</td></tr>'}
async function registrarPagamentoCliente(vendaId,saldo){const valorTexto=prompt(`Saldo atual: ${money(saldo)}
Informe o valor recebido:`);if(valorTexto===null)return;const valor=parseMoney(valorTexto);if(valor<=0)return alert('Informe um valor válido.');if(valor>saldo+0.01)return alert('O valor é maior que o saldo devedor.');const observacao=prompt('Observação do pagamento (opcional):')||'';try{await api('registrarRecebimento',{vendaId,valor,observacao,dataPagamento:hoje()});await carregarReceber();await carregarVendas();await carregarDashboard();await carregarComissoes();alert('Pagamento abatido com sucesso.')}catch(e){alert(e.message)}}
$('#receberBody')?.addEventListener('click',e=>{const p=e.target.closest('.js-receber-pagamento');if(p){registrarPagamentoCliente(p.dataset.venda,Number(p.dataset.saldo||0));return}const r=e.target.closest('.js-recibo-receber');if(r){const venda=state.vendas.find(v=>String(v.id)===String(r.dataset.venda));if(venda)abrirRecibo(venda)}});
$('#buscaReceber')?.addEventListener('input',renderReceber);$('#recarregarReceber')?.addEventListener('click',carregarReceber);
function badgeStatusComissao(status){
  const s=String(status||'Pendente');
  if(s==='Aprovada')return '<span class="commission-status is-approved"><i class="fa-solid fa-circle-check"></i> Aprovada</span>';
  if(s==='Recusada')return '<span class="commission-status is-rejected" title="Venda feita pelo administrador"><i class="fa-solid fa-circle-xmark"></i> Recusada</span>';
  return '<span class="commission-status is-pending"><i class="fa-solid fa-clock"></i> Pendente</span>';
}
async function carregarComissoes(){
  try{
    const j=await api('listarComissoes');
    state.comissoes=j.itens||[];
    $('#comissaoLiberada').textContent=money(j.totalLiberado||0);
    $('#comissaoPendente').textContent=money(j.totalPendente||0);
    $('#comissaoPaga').textContent=money(j.totalPago||0);
    $('#comissaoSaldo').textContent=money(j.saldoPagar||0);
    const tbody=$('#comissoesBody');
    tbody.innerHTML=state.comissoes.map(v=>{
      const status=String(v.statusComissao||'Pendente');
      const acoes=isAdmin()?`<div class="commission-actions">${status!=='Aprovada'?`<button type="button" class="table-action commission-approve" onclick="alterarStatusComissao('${esc(v.id)}','Aprovada')"><i class="fa-solid fa-check"></i> Aprovar</button>`:''}${status!=='Recusada'?`<button type="button" class="table-action commission-reject" onclick="alterarStatusComissao('${esc(v.id)}','Recusada')"><i class="fa-solid fa-xmark"></i> Recusar</button>`:''}</div>`:'';
      return `<tr><td>${esc(v.data)}</td><td>${esc(v.produto)}</td><td>${esc(v.cliente)}</td><td>${money(v.valor)}</td><td><strong>${money(v.comissaoGerada)}</strong></td><td>${badgeStatusComissao(status)}</td>${isAdmin()?`<td>${esc(v.vendedora)}</td><td>${acoes||'—'}</td>`:''}</tr>`;
    }).join('')||`<tr><td colspan="${isAdmin()?8:6}">Nenhuma comissão encontrada.</td></tr>`;
    const resumo=$('#resumoComissoesBody');
    if(resumo)resumo.innerHTML=(j.resumoVendedores||[]).map(x=>`<tr><td>${esc(x.vendedora)}</td><td>${money(x.liberada)}</td><td>${money(x.pendente)}</td><td>${money(x.paga)}</td><td><strong>${money(x.saldoPagar)}</strong></td><td>${x.saldoPagar>0?`<button type="button" class="table-action action-pay" onclick="pagarComissao('${String(x.vendedora).replace(/'/g,"\\'")}',${Number(x.saldoPagar)||0})">Pagar comissão</button>`:'Sem valor aprovado'}</td></tr>`).join('')||'<tr><td colspan="6">Nenhuma comissão aprovada para pagamento.</td></tr>';
    const historico=$('#historicoComissoesBody');
    if(historico)historico.innerHTML=(j.pagamentos||[]).map(p=>`<tr><td>${esc(p.data)}</td><td>${esc(p.vendedora)}</td><td><strong>${money(p.valor)}</strong></td><td>${esc(p.referencia||'Pagamento de comissão')}</td><td>${esc(p.registradoPor||'—')}</td></tr>`).join('')||'<tr><td colspan="5">Nenhum pagamento de comissão registrado.</td></tr>';
  }catch(e){console.error(e)}
}
async function alterarStatusComissao(vendaId,status){
  const acao=status==='Aprovada'?'aprovar':'recusar';
  if(!confirm(`Deseja ${acao} esta comissão?`))return;
  try{await api('alterarStatusComissao',{vendaId,status});await carregarComissoes()}catch(e){alert(e.message)}
}
window.alterarStatusComissao=alterarStatusComissao;
async function pagarComissao(vendedora,disponivel){
  const informado=prompt(`Valor da comissão para ${vendedora}:`,Number(disponivel).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}));
  if(informado===null)return;
  const valor=parseMoney(informado);if(!valor||valor<=0)return alert('Informe um valor válido.');
  try{await api('registrarPagamentoComissao',{vendedora,valor,referencia:'Pagamento registrado pelo sistema'});alert('Pagamento da comissão registrado.');await carregarComissoes()}catch(e){alert(e.message)}
}
$('#recarregarComissoes').onclick=carregarComissoes;

function showApp(){ $('#login').classList.add('hidden');$('#app').classList.remove('hidden');$('#userInfo').textContent=`${state.user.nome} • ${state.user.perfil}`;$$('.admin-only').forEach(x=>x.classList.toggle('hidden',!isAdmin()));$$('.reseller-only').forEach(x=>x.classList.toggle('hidden',!isRevendedor()));if(isRevendedor()){$$('aside .nav:not(.reseller-only)').forEach(x=>x.classList.add('hidden'));$$('.nav').forEach(x=>x.classList.remove('active'));const b=$('[data-page=revendedor]');b.classList.remove('hidden');b.classList.add('active');$$('.page').forEach(x=>x.classList.remove('active'));$('#page-revendedor').classList.add('active');$('#titulo').textContent='Catálogo de revenda';carregarCatalogoRevendedor()}else{carregarTudo();restaurarPaginaAtiva()}}
$('#loginForm').onsubmit=async e=>{e.preventDefault();const d=formObj(e.currentTarget);$('#loginMsg').textContent='Entrando...';try{const j=await api('login',d);state.token=j.token;state.user=j.usuario;localStorage.setItem('dge_token',state.token);showApp()}catch(err){$('#loginMsg').textContent=err.message}}
$('#sair').onclick=()=>{localStorage.removeItem('dge_token');location.reload()};$('#menu').onclick=()=>document.querySelector('aside').classList.toggle('open');
$$('.nav').forEach(b=>b.onclick=()=>{if(b.classList.contains('admin-only')&&!isAdmin())return;$$('.nav').forEach(x=>x.classList.remove('active'));b.classList.add('active');$$('.page').forEach(x=>x.classList.remove('active'));$(`#page-${b.dataset.page}`).classList.add('active');$('#titulo').textContent=b.textContent.trim();document.querySelector('aside').classList.remove('open')});
async function carregarTudo(){if(isRevendedor())return carregarCatalogoRevendedor();await carregarProdutos();await Promise.all([carregarDashboard(),carregarVendas(),carregarReceber(),carregarComissoes(),carregarHorasExtras(),isAdmin()?carregarUsuarios():Promise.resolve()]);const d=$('#vendaForm [name=dataVenda]');if(d&&!d.value)d.value=hoje();const he=$('#horaExtraForm [name=data]');if(he&&!he.value)he.value=hoje()}
async function carregarDashboard(){try{const j=await api('dashboard');$('#fatHoje').textContent=money(j.faturamentoHoje);$('#fatMes').textContent=money(j.faturamentoMes);$('#vendasMes').textContent=j.vendasMes;$('#estoque').textContent=j.estoque;$('#custoMes').textContent=money(j.custoMes);$('#lucroMes').textContent=money(j.lucroMes);$('#dashboardReceber').textContent=money(j.totalReceber||0);$('#dashVendas').innerHTML=(j.ultimasVendas||[]).map(v=>`<tr><td>${esc(v.data)}</td><td>${esc(v.cliente)}</td><td>${esc(v.produto)}</td><td>${money(v.valor)}</td><td>${esc(v.vendedora)}</td></tr>`).join('')||'<tr><td colspan="5">Nenhuma venda.</td></tr>'}catch(e){console.error(e)}}
function promocaoAtiva(v){return ['sim','true','1','yes','x'].includes(String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase())}

function atualizarResumoEstoque(produtos){
  const lista=(produtos||[]).filter(p=>String(p.status||'')!=='Oculto');
  let custoTotal=0,valorVendaTotal=0,quantidadeTotal=0,produtosComPreco=0;
  lista.forEach(p=>{
    const qtd=Math.max(0,Number(p.quantidade)||0);
    const preco=Number(p.preco)||0;
    const custo=Number(p.precoCusto)||0;
    custoTotal+=custo*qtd;
    valorVendaTotal+=preco*qtd;
    quantidadeTotal+=qtd;
    if(preco>0)produtosComPreco++;
  });
  const lucroPotencial=valorVendaTotal-custoTotal;
  const margem=valorVendaTotal>0?(lucroPotencial/valorVendaTotal)*100:0;
  const ticket=quantidadeTotal>0?valorVendaTotal/quantidadeTotal:0;
  const custoEl=$('#custoEstoque'),valorVendaEl=$('#valorVendaEstoque'),lucroEl=$('#lucroEstoque'),margemEl=$('#margemEstoque'),ticketEl=$('#ticketProdutos');
  if(custoEl)custoEl.textContent=money(custoTotal);
  if(valorVendaEl)valorVendaEl.textContent=money(valorVendaTotal);
  if(lucroEl){lucroEl.textContent=money(lucroPotencial);lucroEl.classList.toggle('valor-negativo',lucroPotencial<0);lucroEl.classList.toggle('valor-positivo',lucroPotencial>=0)}
  if(margemEl){margemEl.textContent=margem.toFixed(1).replace('.',',')+'%';margemEl.classList.toggle('valor-negativo',margem<0);margemEl.classList.toggle('valor-positivo',margem>=0)}
  if(ticketEl)ticketEl.textContent=money(ticket);
}

async function carregarProdutos(){try{const j=await api('listarProdutos');state.produtos=j.produtos||[];atualizarResumoEstoque(state.produtos);$('#produtoVenda').innerHTML='<option value="">Selecione</option>'+state.produtos.filter(p=>p.status==='Publicado'&&Number(p.quantidade)>0).map(p=>`<option value="${esc(p.id||('LINHA-'+p.linha))}" data-linha="${esc(p.linha||'')}" data-preco="${p.preco}">${esc(p.nome)} • ${esc(p.categoria)} • ${money(p.preco)} • estoque ${p.quantidade}</option>`).join('');$('#produtosBody').innerHTML=state.produtos.map(p=>{const ativa=promocaoAtiva(p.promocao),reservado=p.status==='Reservado';return `<tr><td>${p.fotos?.[0]?`<img class="thumb" src="${esc(p.fotos[0])}">`:'Sem foto'}</td><td><strong>${esc(p.nome)}</strong><br><small>${esc(p.descricao||'')}</small></td><td>${esc(p.categoria)}</td><td class='preco-cell'>${money(p.preco)}</td><td class='revenda-cell'>${money(p.precoRevenda||calcularPrecoRevendaFront(p.preco,p.categoria))}</td><td class='custo-cell'>${p.precoCusto?money(p.precoCusto):'—'}</td><td class='lucro-cell ${p.precoCusto?(((+p.preco)-(+p.precoCusto))>=0?'lucro-positivo':'lucro-negativo'):''}'>${p.precoCusto?money((+p.preco)-(+p.precoCusto)):'—'}</td><td class='margem-cell ${p.precoCusto?(((+p.preco)-(+p.precoCusto))>=0?'margem-positiva':'margem-negativa'):''}'>${p.precoCusto&&(+p.preco)?((((+p.preco)-(+p.precoCusto))/(+p.preco))*100).toFixed(1)+'%':'—'}</td><td>${p.quantidade}</td><td><span class="pill">${esc(p.status)}</span></td><td><div class="row-actions"><button type="button" class="action-btn action-edit" onclick="editarProduto('${p.id}')"><i class="fa-solid fa-pen"></i><span>Editar</span></button>${podeGerenciarProdutos()?`<button type="button" class="action-btn ${ativa?'promo-active':'promo-btn'}" onclick="promocaoProduto('${p.id}',${ativa?'false':'true'},'${p.linha||''}',this)"><i class="fa-solid fa-tag"></i><span>${ativa?'Remover promoção':'Promoção'}</span></button><button type="button" class="action-btn action-reserve" onclick="statusProduto('${p.id}','${reservado?'Publicado':'Reservado'}','${p.linha||''}')"><i class="fa-solid fa-${reservado?'unlock':'lock'}"></i><span>${reservado?'Liberar':'Reservar'}</span></button><button type="button" class="action-btn action-visibility" onclick="statusProduto('${p.id}','${p.status==='Oculto'?'Publicado':'Oculto'}','${p.linha||''}')"><i class="fa-solid fa-${p.status==='Oculto'?'eye':'eye-slash'}"></i><span>${p.status==='Oculto'?'Publicar':'Ocultar'}</span></button><button type="button" class="action-btn danger" onclick="excluirProduto('${p.id}')"><i class="fa-solid fa-trash"></i><span>Excluir</span></button>`:''}</div></td></tr>`}).join('')||'<tr><td colspan="11">Nenhum produto.</td></tr>'}catch(e){console.error(e);alert('Erro ao carregar produtos: '+e.message)}}
async function carregarVendas(){try{const j=await api('listarVendas');state.vendas=j.vendas||[];$('#vendasBody').innerHTML=state.vendas.map(v=>`<tr><td>${esc(v.data)}</td><td>${esc(v.cliente)}<br><small>${esc(v.documento||'')}</small></td><td>${esc(v.produto)}</td><td>${money(v.valor)}${v.saldo>0?`<br><small class="saldo-pendente">Falta ${money(v.saldo)}</small>`:''}</td><td>${esc(v.pagamento)}${v.pagamento==='No prazo'?`<br><small>${v.parcelasPagas||0}/${v.totalParcelas||0} pagas</small>`:''}</td><td>${esc(v.vendedora)}</td><td><button class="table-action action-receipt" onclick="imprimirRecibo('${v.id}')">Recibo</button>${isAdmin()?` <button class="table-action action-delete" onclick="excluirVenda('${v.id}')">Excluir</button>`:''}</td></tr>`).join('')||'<tr><td colspan="7">Nenhuma venda.</td></tr>'}catch(e){console.error(e)}}
$('#recarregarVendas').onclick=carregarVendas;


$('#vendaForm').onsubmit=async e=>{e.preventDefault();const f=e.currentTarget,d=formObj(f);const msg=f.querySelector('.msg');const botao=f.querySelector('button[type=submit],button.primary');if(!validarDocumento(d.documento)){msg.textContent='CPF ou CNPJ inválido.';f.documento.focus();return}if(!validarTelefone(d.telefone)){msg.textContent='Telefone inválido. Informe DDD e número.';f.telefone.focus();return}if(d.tipoVenda==='avulsa'&&!String(d.produtoManual||'').trim()){msg.textContent='Informe o produto não cadastrado.';f.produtoManual.focus();return}if(d.tipoVenda!=='avulsa'&&!String(d.idProduto||'').trim()){msg.textContent='Selecione um produto do estoque.';f.idProduto.focus();return}d.documento=somenteDigitos(d.documento);d.telefone=somenteDigitos(d.telefone);d.valorVenda=parseMoney(d.valorVenda);d.quantidade=Number(d.quantidade||1);if(d.valorVenda<=0){msg.textContent='Informe um valor de venda válido.';f.valorVenda.focus();return}if(d.pagamento==='No prazo'){d.entrada=parseMoney(d.entrada);if(d.entrada<0||d.entrada>d.valorVenda){msg.textContent='A entrada não pode ser maior que o valor da venda.';return}}else{d.entrada=d.valorVenda}d.parcelas=[];msg.textContent='Salvando...';if(botao)botao.disabled=true;try{const j=await api('cadastrarVenda',d);msg.textContent='Venda salva.';if(d.imprimir)abrirRecibo(j.venda);f.reset();f.dataVenda.value=hoje();f.tipoVenda.value='estoque';alternarTipoVenda();alternarPagamento();await carregarTudo()}catch(err){msg.textContent='Erro: '+err.message}finally{if(botao)botao.disabled=false}}
function abrirRecibo(v){
  const w=open('','_blank','width=920,height=980');
  const quantidade=Number(v.quantidade||1),valorTotal=Number(v.valor||0),valorUnitario=quantidade>0?valorTotal/quantidade:valorTotal;
  const garantia=Number(v.garantia||0),garantiaTexto=garantia>0?`${garantia} dias`:'Sem garantia informada';
  const parcelas=Array.isArray(v.parcelas)?v.parcelas:[],prazo=(v.pagamento==='No prazo'||parcelas.length>0||Number(v.saldo||0)>0);
  const entrada=Number(v.entrada||0),saldo=Number(v.saldo||0);
  const fmtDataRecibo=data=>{if(!data)return'—';const s=String(data);if(/^\d{4}-\d{2}-\d{2}$/.test(s)){const [a,m,d]=s.split('-');return`${d}/${m}/${a}`}return s};
  const linhasParcelas=parcelas.map((p,i)=>`<tr><td class="center">${Number(p.numero||i+1)}ª</td><td>${fmtDataRecibo(p.vencimento)}</td><td class="right">${money(p.valor)}</td><td class="center"><span class="parcel-status ${p.status==='Pago'?'ok':'pending'}">${esc(p.status||'Pendente')}</span></td></tr>`).join('');
  const blocoParcelas=prazo?`<section class="section"><div class="section-title">Plano de pagamento — ${parcelas.length} parcela${parcelas.length===1?'':'s'}</div><div class="credit-summary"><div><span>Valor da venda</span><strong>${money(valorTotal)}</strong></div><div><span>Entrada recebida</span><strong>${money(entrada)}</strong></div><div><span>Saldo a receber</span><strong>${money(saldo)}</strong></div></div>${parcelas.length?`<table class="product-table installment-table"><thead><tr><th class="center">Parcela</th><th>Data de vencimento</th><th class="right">Valor</th><th class="center">Status</th></tr></thead><tbody>${linhasParcelas}</tbody></table>`:'<div class="notes">Venda a prazo registrada sem calendário de parcelas.</div>'}</section>`:'';
  const declaracao=prazo?`Declaramos que a venda acima foi realizada a prazo. A entrada de <strong>${money(entrada)}</strong> foi registrada, permanecendo o saldo de <strong>${money(saldo)}</strong>, a ser pago conforme as parcelas e datas de vencimento descritas neste documento.`:`Declaramos que recebemos do cliente acima identificado o valor total informado, referente à aquisição do produto descrito neste documento.`;
  w.document.write(`<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Recibo ${esc(v.id)}</title><style>
  *{box-sizing:border-box}body{margin:0;background:#eef0f3;color:#171717;font-family:Arial,Helvetica,sans-serif}.page{width:210mm;min-height:297mm;margin:20px auto;background:#fff;padding:17mm 16mm 15mm;box-shadow:0 12px 35px rgba(0,0,0,.12);position:relative}.header{display:flex;justify-content:space-between;gap:28px;align-items:flex-start;padding-bottom:20px;border-bottom:3px solid #f5b400}.brand{display:flex;align-items:center;gap:14px}.brand-mark{width:58px;height:58px;border-radius:14px;background:#111;color:#f5b400;display:grid;place-items:center;font-size:25px;font-weight:900}.brand-name{font-size:27px;font-weight:900}.brand-name span{color:#d89c00}.company{margin-top:9px;color:#555;font-size:12px;line-height:1.55}.receipt-head{text-align:right}.receipt-head h1{font-size:22px;margin:0 0 8px}.receipt-number{display:inline-block;background:#111;color:#fff;border-radius:999px;padding:7px 12px;font-size:11px;font-weight:700}.paid{display:inline-block;margin-top:12px;border:1px solid ${prazo?'#d89c00':'#1f9d57'};color:${prazo?'#775700':'#14723e'};background:${prazo?'#fff7d6':'#e8f8ef'};border-radius:8px;padding:7px 11px;font-size:11px;font-weight:800}.section{margin-top:24px}.section-title{font-size:11px;font-weight:900;letter-spacing:1.2px;text-transform:uppercase;color:#777;margin-bottom:9px}.customer-grid{display:grid;grid-template-columns:2fr 1fr 1fr;border:1px solid #e1e1e1;border-radius:12px;overflow:hidden}.field{padding:13px 15px;border-right:1px solid #e8e8e8}.field:last-child{border-right:0}.label{display:block;font-size:10px;text-transform:uppercase;color:#888;margin-bottom:4px;font-weight:700}.value{font-size:13px;font-weight:700}.product-table{width:100%;border-collapse:separate;border-spacing:0;border:1px solid #dedede;border-radius:12px;overflow:hidden}.product-table th{background:#f6f6f6;text-transform:uppercase;font-size:10px;color:#666;text-align:left;padding:11px 13px;border-bottom:1px solid #dedede}.product-table td{padding:13px;font-size:12px;border-bottom:1px solid #eee}.product-table tr:last-child td{border-bottom:0}.right{text-align:right!important}.center{text-align:center!important}.summary{margin-left:auto;width:330px;margin-top:14px;border:1px solid #dedede;border-radius:12px;overflow:hidden}.summary-row{display:flex;justify-content:space-between;padding:10px 13px;border-bottom:1px solid #eee;font-size:12px}.summary-row.total{background:#111;color:#fff;font-size:18px;font-weight:900}.summary-row.total strong{color:#f5b400}.credit-summary{display:grid;grid-template-columns:repeat(3,1fr);border:1px solid #dedede;border-radius:12px;overflow:hidden;margin-bottom:12px}.credit-summary div{padding:13px;border-right:1px solid #eee}.credit-summary div:last-child{border-right:0}.credit-summary span{display:block;font-size:10px;text-transform:uppercase;color:#777;margin-bottom:5px}.credit-summary strong{font-size:15px}.parcel-status{display:inline-block;padding:4px 8px;border-radius:999px;font-size:10px;font-weight:800}.parcel-status.ok{background:#e8f8ef;color:#14723e}.parcel-status.pending{background:#fff2cc;color:#775700}.notes{border:1px solid #e1e1e1;background:#fafafa;border-radius:12px;padding:14px 15px;min-height:56px;font-size:12px}.declaration{margin-top:22px;font-size:11px;line-height:1.6;color:#555;text-align:justify}.signatures{display:grid;grid-template-columns:1fr 1fr;gap:45px;margin-top:56px}.signature{text-align:center}.signature-line{border-top:1px solid #333;padding-top:8px;font-size:12px;font-weight:800}.footer{margin-top:36px;border-top:1px solid #e4e4e4;padding-top:8px;display:flex;justify-content:space-between;color:#888;font-size:9px}.actions{width:210mm;margin:0 auto 30px;display:flex;justify-content:flex-end;gap:10px}.actions button{border:0;border-radius:9px;padding:12px 18px;font-weight:800;cursor:pointer}.print{background:#f5b400}.close{background:#222;color:#fff}@media(max-width:850px){body{background:#fff}.page{width:100%;min-height:auto;margin:0;padding:24px;box-shadow:none}.customer-grid,.credit-summary{grid-template-columns:1fr}.field,.credit-summary div{border-right:0;border-bottom:1px solid #e8e8e8}.summary{width:100%}.signatures{grid-template-columns:1fr}.actions{width:100%;padding:0 24px}}@media print{@page{size:A4;margin:0}body{background:#fff}.page{margin:0;width:210mm;min-height:297mm;box-shadow:none}.actions{display:none}}
  </style></head><body><main class="page"><header class="header"><div class="brand"><div class="brand-mark">DG</div><div><div class="brand-name">DG <span>ELETROS</span></div><div class="company"><strong>DG ELETROS LTDA ME</strong><br>CNPJ 60.364.662/0001-72<br>Rua Afonso Vidal, 264 — Vila Andrade<br>São Paulo/SP — CEP 05723-330<br>Telefone: (11) 97971-8211</div></div></div><div class="receipt-head"><h1>${prazo?'RECIBO DE VENDA A PRAZO':'RECIBO DE VENDA'}</h1><div class="receipt-number">Nº ${esc(v.id)}</div><br><div class="paid">${prazo?'PAGAMENTO PARCELADO':'PAGAMENTO REGISTRADO'}</div></div></header><section class="section"><div class="section-title">Dados do cliente</div><div class="customer-grid"><div class="field"><span class="label">Nome / Razão social</span><span class="value">${esc(v.cliente)}</span></div><div class="field"><span class="label">CPF / CNPJ</span><span class="value">${esc(v.documento||'Não informado')}</span></div><div class="field"><span class="label">Telefone</span><span class="value">${esc(v.telefone||'Não informado')}</span></div></div></section><section class="section"><div class="section-title">Itens da venda</div><table class="product-table"><thead><tr><th>Descrição do produto</th><th class="center">Qtd.</th><th class="right">Valor unitário</th><th class="right">Subtotal</th></tr></thead><tbody><tr><td><strong>${esc(v.produto)}</strong></td><td class="center">${quantidade}</td><td class="right">${money(valorUnitario)}</td><td class="right"><strong>${money(valorTotal)}</strong></td></tr></tbody></table><div class="summary"><div class="summary-row"><span>Forma de pagamento</span><strong>${esc(v.pagamento||'Não informado')}</strong></div><div class="summary-row"><span>Data da venda</span><strong>${esc(v.data)}</strong></div><div class="summary-row"><span>Garantia</span><strong>${garantiaTexto}</strong></div><div class="summary-row total"><span>${prazo?'Valor da venda':'Total pago'}</span><strong>${money(valorTotal)}</strong></div></div></section>${blocoParcelas}<section class="section"><div class="section-title">Observações</div><div class="notes">${esc(v.observacoes||'Nenhuma observação registrada.')}</div></section><p class="declaration">${declaracao}</p><div class="signatures"><div class="signature"><div class="signature-line">DG ELETROS LTDA ME</div></div><div class="signature"><div class="signature-line">${esc(v.cliente)}</div></div></div><footer class="footer"><span>Documento gerado pelo sistema DG Eletros</span><span>Recibo nº ${esc(v.id)}</span></footer></main><div class="actions"><button class="close" onclick="window.close()">Fechar</button><button class="print" onclick="window.print()">Imprimir / Salvar PDF</button></div></body></html>`);w.document.close();
}
window.imprimirRecibo=id=>{const v=state.vendas.find(x=>x.id===id);if(v)abrirRecibo(v)}
window.excluirVenda=async id=>{if(!confirm('Excluir esta venda e devolver a quantidade ao estoque?'))return;const motivo=prompt('Motivo da exclusão:');if(!motivo)return;try{await api('excluirVenda',{id,motivo});await carregarTudo()}catch(e){alert(e.message)}}
function abrirModal(html){$('#modalConteudo').innerHTML=html;$('#modal').classList.remove('hidden')}$('#fecharModal').onclick=()=>$('#modal').classList.add('hidden');$('#modal').onclick=e=>{if(e.target===$('#modal'))$('#modal').classList.add('hidden')}
$('#novoProduto').onclick=()=>formProduto();window.editarProduto=id=>formProduto(state.produtos.find(p=>p.id===id));
function formProduto(p={}){abrirModal(`<h2>${p.id?'Editar':'Novo'} produto</h2><form id="produtoForm" class="grid"><input type="hidden" name="id" value="${esc(p.id||'')}"><input type="hidden" name="linha" value="${esc(p.linha||'')}"><label>Nome do produto<input name="nome" required value="${esc(p.nome||'')}"></label><label>Categoria<select name="categoria">${['Geladeira','Lavadora','Fogão','Micro-ondas','TV','Ar-condicionado','Ferramenta','Portátil','Eletrônico','Outro'].map(x=>`<option ${p.categoria===x?'selected':''}>${x}</option>`).join('')}</select></label><label>Marca<input name="marca" value="${esc(p.marca||'')}"></label><label>Modelo<input name="modelo" value="${esc(p.modelo||'')}"></label><label>Classificação<select name="classificacao"><option ${p.classificacao==='A'?'selected':''}>A</option><option ${p.classificacao==='B'?'selected':''}>B</option></select></label><label>Preço no PIX<input name="preco" required value="${p.preco?Number(p.preco).toLocaleString('pt-BR',{minimumFractionDigits:2}):''}"></label><label>Preço para revendedor<input value="${money(p.precoRevenda||calcularPrecoRevendaFront(p.preco,p.categoria))}" readonly><small>Calculado automaticamente: 11% linha pesada e 18% linha leve.</small></label><label>Preço de custo (opcional)<input name="precoCusto" value="${p.precoCusto?Number(p.precoCusto).toLocaleString('pt-BR',{minimumFractionDigits:2}):''}" placeholder="Quanto foi pago no produto"></label><label>Quantidade<input type="number" name="quantidade" min="0" required value="${p.quantidade??1}"></label><label>Status<select name="status">${['Publicado','Oculto','Reservado'].map(x=>`<option ${p.status===x?'selected':''}>${x}</option>`).join('')}</select></label><label class="full">Descrição<textarea name="descricao" rows="3">${esc(p.descricao||'')}</textarea></label><label class="full">Comentário exclusivo para revendedor<textarea name="comentarioRevendedor" rows="3" placeholder="Informe avarias, peça trocada ou observação importante">${esc(p.comentarioRevendedor||'')}</textarea></label><label class="full">Fotos do produto<input type="file" name="fotos" accept="image/*" multiple><small>Até 6 fotos. Ao editar, novas fotos substituem as atuais.</small></label><div class="full preview">${(p.fotos||[]).map(x=>`<img src="${esc(x)}">`).join('')}</div><div class="full actions"><button class="primary">Salvar produto</button><span class="msg"></span></div></form>`);$('#produtoForm').onsubmit=salvarProduto}
async function fileData(file){const data=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(file)});return {nome:file.name,data}}
async function salvarProduto(e){e.preventDefault();const f=e.currentTarget,d=formObj(f);d.preco=parseMoney(d.preco);d.precoCusto=parseMoney(d.precoCusto);d.quantidade=Number(d.quantidade||0);const files=[...f.fotos.files];if(files.length>6)return alert('Selecione no máximo 6 fotos.');if(files.length){d.fotos=[];for(const file of files)d.fotos.push(await fileData(file))}delete d.fotos;d.fotosUpload=[];for(const file of files)d.fotosUpload.push(await fileData(file));const msg=f.querySelector('.msg');msg.textContent='Salvando...';try{await api('salvarProduto',d);$('#modal').classList.add('hidden');await carregarTudo()}catch(err){msg.textContent=err.message}}
window.statusProduto=async(id,status,linha)=>{try{const j=await api('alterarStatusProduto',{id,status,linha:Number(linha||0)});alert(j.status==='Reservado'?'Produto marcado como reservado.':j.status==='Oculto'?'Produto ocultado do site.':'Produto liberado e publicado.');await carregarProdutos()}catch(e){alert(e.message)}}
window.promocaoProduto=async(id,ativar,linha,botao)=>{const texto=ativar?'colocar este produto na promoção da semana?':'remover este produto da promoção da semana?';if(!confirm('Deseja '+texto))return;const anterior=botao?.textContent;if(botao){botao.disabled=true;botao.textContent='Salvando...'}try{await api('alterarPromocaoProduto',{id,linha:Number(linha||0),promocao:ativar?'Sim':'Não'});await carregarProdutos();alert(ativar?'Produto colocado na promoção da semana.':'Produto removido da promoção.')}catch(e){if(botao){botao.disabled=false;botao.textContent=anterior}alert('Não foi possível alterar a promoção: '+e.message)}}

window.excluirProduto=async id=>{if(!confirm('Excluir definitivamente este produto?'))return;const motivo=prompt('Motivo da exclusão:');if(!motivo)return;try{await api('excluirProduto',{id,motivo});await carregarTudo()}catch(e){alert(e.message)}}
async function carregarUsuarios(){
  try{
    const j=await api('listarUsuarios');
    state.usuarios=j.usuarios||[];
    $('#usuariosBody').innerHTML=state.usuarios.map(u=>`<tr>
      <td><div class="employee-name"><span class="employee-avatar">${esc((u.nome||'U').charAt(0).toUpperCase())}</span><strong>${esc(u.nome)}</strong></div></td>
      <td><span class="employee-user">@${esc(u.usuario)}</span></td>
      <td><span class="employee-role">${esc(u.perfil)}</span></td>
      <td>${Number(u.salario||0)>0?`<strong>${money(u.salario)}</strong>`:'—'}</td>
      <td>${Number(u.horasSemanais||0)>0?`${Number(u.horasSemanais).toLocaleString('pt-BR')}h`:'—'}</td>
      <td><span class="employee-status ${String(u.status).toLowerCase()==='ativo'?'is-active':'is-inactive'}"><i class="fa-solid fa-circle"></i>${esc(u.status)}</span></td>
      <td><div class="employee-actions"><button type="button" class="employee-btn employee-edit" onclick="editarUsuario('${u.id}')"><i class="fa-solid fa-pen-to-square"></i><span>Editar</span></button><button type="button" class="employee-btn employee-delete" onclick="excluirUsuario('${u.id}')"><i class="fa-solid fa-trash-can"></i><span>Excluir</span></button></div></td>
    </tr>`).join('')||'<tr><td colspan="7">Nenhum funcionário cadastrado.</td></tr>';
  }catch(e){console.error(e)}
}
$('#novoUsuario').onclick=()=>formUsuario();window.editarUsuario=id=>formUsuario(state.usuarios.find(u=>u.id===id));
function formUsuario(u={}){
  abrirModal(`<h2>${u.id?'Editar':'Novo'} funcionário</h2><form id="usuarioForm" class="grid">
    <input type="hidden" name="id" value="${esc(u.id||'')}">
    <label>Nome<input name="nome" required value="${esc(u.nome||'')}"></label>
    <label>Usuário<input name="usuario" required value="${esc(u.usuario||'')}"></label>
    <label>Senha<input name="senha" type="password" ${u.id?'':'required'} placeholder="${u.id?'Deixe em branco para manter':''}"></label>
    <label>Perfil<select name="perfil"><option ${u.perfil==='Vendedora'?'selected':''}>Vendedora</option><option ${u.perfil==='Administrador'?'selected':''}>Administrador</option><option ${u.perfil==='Revendedor'?'selected':''}>Revendedor</option></select></label>
    <label>Salário mensal (R$)<input name="salario" inputmode="decimal" value="${Number(u.salario||0)>0?Number(u.salario).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}):''}" placeholder="Ex.: 2.100,00"><small>Usado para calcular o valor normal da hora.</small></label>
    <label>Carga horária semanal<input name="horasSemanais" type="number" min="1" max="60" step="0.5" value="${u.horasSemanais||''}" placeholder="Ex.: 44"><small>Ex.: 44h, 40h ou 36h por semana.</small></label>
    <label>Status<select name="status"><option ${u.status!=='Inativo'?'selected':''}>Ativo</option><option ${u.status==='Inativo'?'selected':''}>Inativo</option></select></label>
    <div class="full employee-hour-info"><i class="fa-solid fa-calculator"></i><span>Hora normal = salário ÷ (horas semanais × 5). Hora extra = valor da hora normal + 50%.</span></div>
    <div class="full"><button class="primary">Salvar</button></div>
  </form>`);
  $('#usuarioForm').onsubmit=async e=>{
    e.preventDefault();
    const d=formObj(e.currentTarget);
    d.salario=parseMoney(d.salario);
    d.horasSemanais=Number(d.horasSemanais||0);
    if(d.perfil==='Vendedora'&&(d.salario<=0||d.horasSemanais<=0))return alert('Informe o salário e a carga horária semanal da vendedora.');
    try{await api('salvarUsuario',d);$('#modal').classList.add('hidden');await carregarUsuarios();await carregarHorasExtras()}catch(err){alert(err.message)}
  };
}

function formatarMinutosExtras(total){
  total=Math.max(0,Number(total)||0);
  const h=Math.floor(total/60),m=Math.round(total%60);
  return `${h}h${m?` ${m}min`:''}`;
}
async function carregarHorasExtras(){
  try{
    const j=await api('listarHorasExtras');
    state.horasExtras=j.registros||[];
    if($('#heHorasPendentes'))$('#heHorasPendentes').textContent=formatarMinutosExtras(j.minutosPendentes||0);
    if($('#heValorPendente'))$('#heValorPendente').textContent=money((Number(j.valorPendente||0))*1.5);
    if($('#heHorasPagas'))$('#heHorasPagas').textContent=formatarMinutosExtras(j.minutosPagos||0);
    if($('#heValorTotal'))$('#heValorTotal').textContent=money((Number(j.valorTotal||0))*1.5);
    const tbody=$('#horasExtrasBody');if(!tbody)return;
    tbody.innerHTML=state.horasExtras.map(r=>{
      const pago=String(r.status||'Pendente')==='Pago',acoes=[];
      if(isAdmin())acoes.push(`<button type="button" class="employee-btn ${pago?'employee-edit':'overtime-pay'}" onclick="alterarStatusHoraExtra('${r.id}','${pago?'Pendente':'Pago'}')"><i class="fa-solid fa-${pago?'rotate-left':'check'}"></i><span>${pago?'Reabrir':'Marcar paga'}</span></button>`);
      if(isAdmin()||String(r.idFuncionario||'')===String(state.user?.id||''))acoes.push(`<button type="button" class="employee-btn employee-delete" onclick="excluirHoraExtra('${r.id}')"><i class="fa-solid fa-trash-can"></i><span>Excluir</span></button>`);
      return `<tr><td><strong>${esc(r.data||'')}</strong></td><td>${esc(r.funcionario||'')}</td><td><strong>${formatarMinutosExtras(r.minutos||0)}</strong></td><td>${esc(r.motivo||'')}</td><td>${money((Number(r.valorHora||0))*1.5)}</td><td><strong>${money((Number(r.valorTotal||0))*1.5)}</strong></td><td><span class="overtime-status ${pago?'is-paid':'is-pending'}"><i class="fa-solid fa-circle"></i>${pago?'Pago':'Pendente'}</span></td><td><div class="employee-actions">${acoes.join('')||'—'}</div></td></tr>`;
    }).join('')||'<tr><td colspan="8">Nenhuma hora extra cadastrada.</td></tr>';
  }catch(e){console.error(e);const tbody=$('#horasExtrasBody');if(tbody)tbody.innerHTML=`<tr><td colspan="8">${esc(e.message)}</td></tr>`}
}
$('#horaExtraForm')?.addEventListener('submit',async e=>{
  e.preventDefault();const f=e.currentTarget,d=formObj(f),msg=$('#horaExtraMsg');d.horas=Math.max(0,Number(d.horas||0));d.minutos=Math.max(0,Number(d.minutos||0));
  if((d.horas*60+d.minutos)<=0)return alert('Informe a quantidade de horas ou minutos extras.');if(msg)msg.textContent='Salvando...';
  try{const j=await api('registrarHoraExtra',d);f.reset();f.querySelector('[name=data]').value=hoje();f.querySelector('[name=horas]').value=0;f.querySelector('[name=minutos]').value=0;if(msg)msg.textContent=`Registrado: ${formatarMinutosExtras(j.registro?.minutos||0)} • ${money((Number(j.registro?.valorTotal||0))*1.5)}`;await carregarHorasExtras()}catch(err){if(msg)msg.textContent=err.message}
});
$('#recarregarHorasExtras')?.addEventListener('click',carregarHorasExtras);
window.alterarStatusHoraExtra=async(id,status)=>{try{await api('alterarStatusHoraExtra',{id,status});await carregarHorasExtras()}catch(e){alert(e.message)}};
window.excluirHoraExtra=async id=>{if(!confirm('Excluir este lançamento de hora extra?'))return;try{await api('excluirHoraExtra',{id});await carregarHorasExtras()}catch(e){alert(e.message)}};

iniciarTema();
protegerSessaoEInterface();
(async()=>{if(!state.token)return;try{const j=await api('verificarToken');state.user=j.usuario;showApp()}catch(e){localStorage.removeItem('dge_token')}})();


function normalizarTexto(v){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()}
function linhaPesadaFront(categoria){const c=normalizarTexto(categoria).replace(/-/g,' ');return ['geladeira','lavadora','fogao','micro ondas','microondas','ar condicionado','frigobar','maquina de lavar','lava e seca','freezer'].some(x=>c===x||c.includes(x))}
function calcularPrecoRevendaFront(preco,categoria){return Math.round((Number(preco||0)*(linhaPesadaFront(categoria)?.89:.82))*100)/100}
let produtosRevenda=[];
async function carregarCatalogoRevendedor(){
  const box=$('#catalogoRevendedor');if(!box)return;box.innerHTML='<div class="catalog-loading">Carregando produtos...</div>';
  try{const j=await api('listarCatalogoRevendedor');produtosRevenda=j.produtos||[];renderCatalogoRevendedor()}catch(e){box.innerHTML=`<div class="catalog-empty">${esc(e.message)}</div>`}
}
function renderCatalogoRevendedor(){
  const busca=normalizarTexto($('#buscaRevendedor')?.value);const ordem=$('#ordemRevendedor')?.value;let lista=produtosRevenda.filter(p=>!busca||normalizarTexto([p.nome,p.marca,p.modelo,p.categoria].join(' ')).includes(busca));
  if(ordem==='desconto')lista.sort((a,b)=>(b.precoNormal-b.precoRevenda)-(a.precoNormal-a.precoRevenda));if(ordem==='menor-preco')lista.sort((a,b)=>a.precoRevenda-b.precoRevenda);if(ordem==='maior-preco')lista.sort((a,b)=>b.precoRevenda-a.precoRevenda);
  $('#totalRevendedor').textContent=lista.length;const box=$('#catalogoRevendedor');box.innerHTML=lista.map((p,idx)=>{const fotos=(p.fotos||[]).filter(Boolean),gid='dge-gallery-'+idx,desconto=Math.max(0,p.precoNormal-p.precoRevenda),msg=encodeURIComponent(`Olá! Tenho interesse no produto ${p.nome}${p.modelo?' - '+p.modelo:''}, pelo preço de revenda de ${money(p.precoRevenda)}.`);return `<article class="reseller-card"><div class="reseller-gallery-wrap" id="${gid}"><div class="reseller-gallery">${fotos.length?fotos.map((f,i)=>`<img src="${esc(f)}" alt="${esc(p.nome)} foto ${i+1}">`).join(''):'<div class="photo-empty"><i class="fa-solid fa-box-open"></i></div>'}</div>${fotos.length>1?`<button class="gallery-arrow prev" onclick="moverGaleria('${gid}',-1)"><i class="fa-solid fa-chevron-left"></i></button><button class="gallery-arrow next" onclick="moverGaleria('${gid}',1)"><i class="fa-solid fa-chevron-right"></i></button><div class="gallery-dots">${fotos.map((_,i)=>`<button class="gallery-dot ${i===0?'active':''}" onclick="irFoto('${gid}',${i})"></button>`).join('')}</div>`:''}<span class="discount-badge">-${p.descontoPercentual}%</span></div><div class="reseller-body"><span class="category-badge">${esc(p.categoria)}</span><h2>${esc(p.nome)}</h2><p class="specs">${esc([p.marca,p.modelo,p.classificacao?('Classificação '+p.classificacao):''].filter(Boolean).join(' • '))}</p>${p.comentarioRevendedor?`<div class="reseller-note"><i class="fa-solid fa-circle-info"></i><div><strong>Informação importante</strong><p>${esc(p.comentarioRevendedor)}</p></div></div>`:''}<div class="price-compare"><div><small>Preço normal</small><strong class="normal-price">${money(p.precoNormal)}</strong></div><div><small>Preço revendedor</small><strong class="resale-price">${money(p.precoRevenda)}</strong></div></div><div class="saving"><i class="fa-solid fa-arrow-trend-down"></i> Economia de ${money(desconto)}</div><a class="whatsapp-interest" target="_blank" href="https://wa.me/5511979718211?text=${msg}"><i class="fa-brands fa-whatsapp"></i> Tenho interesse</a></div></article>`}).join('')||'<div class="catalog-empty"><i class="fa-solid fa-magnifying-glass"></i><strong>Nenhum produto encontrado</strong><p>Tente pesquisar outro modelo ou categoria.</p></div>';
}
window.moverGaleria=(id,dir)=>{const el=document.querySelector('#'+id+' .reseller-gallery');if(el)el.scrollBy({left:el.clientWidth*dir,behavior:'smooth'})};
window.irFoto=(id,i)=>{const el=document.querySelector('#'+id+' .reseller-gallery');if(el)el.scrollTo({left:el.clientWidth*i,behavior:'smooth'})};
$('#buscaRevendedor')?.addEventListener('input',renderCatalogoRevendedor);$('#ordemRevendedor')?.addEventListener('change',renderCatalogoRevendedor);
window.excluirUsuario=async id=>{if(!confirm('Excluir este cadastro?'))return;try{await api('excluirUsuario',{id});carregarUsuarios()}catch(e){alert(e.message)}};
