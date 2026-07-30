const CFG={ABAS:{PRODUTOS:'Produtos',VENDAS:'Vendas',USUARIOS:'Usuarios',LOG:'LogExclusoes'},TOKEN_HORAS:12,PASTA:'DG Eletros - Fotos dos Produtos'};
const PLANILHA_ID='1FCiKWKy_DMa3Dwj12cCD2KM_qYFuTnRlKSQ9fM2FleA';
function planilha_(){return SpreadsheetApp.openById(PLANILHA_ID);}
function configurarSistema(){const ss=planilha_();criarAba_(ss,CFG.ABAS.PRODUTOS,['Produto','Categoria','Preço','Preço de custo','Classificação','Descrição','Link da Imagem','Promoção','WhatsApp','ID','Marca','Modelo','Quantidade','Status','Fotos','Data de cadastro']);criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria']);criarAba_(ss,CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro']);criarAba_(ss,CFG.ABAS.LOG,['Data','Tipo','ID','Motivo','Usuário']);migrarProdutosExistentes_(ss.getSheetByName(CFG.ABAS.PRODUTOS));const sh=ss.getSheetByName(CFG.ABAS.USUARIOS);if(sh.getLastRow()<2)sh.appendRow([uid_('USR'),'Administrador','admin',hash_('123456'),'Administrador','Ativo','','',new Date()]);return 'Sistema configurado. Login inicial: admin / 123456';}
function localizarCabecalho_(sh){
  if(!sh)return 1;
  const maxLin=Math.min(Math.max(sh.getLastRow(),1),50);
  const maxCol=Math.max(sh.getLastColumn(),7);
  const vals=sh.getRange(1,1,maxLin,maxCol).getDisplayValues();
  for(let i=0;i<vals.length;i++){
    const norm=vals[i].map(v=>String(v).trim().toLowerCase());
    if(norm.includes('produto')&&norm.includes('categoria')&&(norm.includes('preço')||norm.includes('preco')))return i+1;
  }
  return 1;
}
function cabecalhos_(sh){
  const linha=localizarCabecalho_(sh);
  const ultima=Math.max(sh.getLastColumn(),1);
  return{linha,cab:sh.getRange(linha,1,1,ultima).getValues()[0]};
}
function criarAba_(ss,nome,cab){
  let sh=ss.getSheetByName(nome);
  if(!sh){sh=ss.insertSheet(nome);sh.getRange(1,1,1,cab.length).setValues([cab]);}
  const info=cabecalhos_(sh);
  let atual=info.cab.map(v=>String(v).trim());
  if(!atual.some(Boolean)){
    sh.getRange(info.linha,1,1,cab.length).setValues([cab]);
    atual=[...cab];
  }else{
    cab.forEach(c=>{
      if(!atual.includes(c)){
        const novaCol=atual.length+1;
        sh.getRange(info.linha,novaCol).setValue(c);
        atual.push(c);
      }
    });
  }
  sh.setFrozenRows(info.linha);
  return sh;
}
function doGet(e){try{const acao=String(e?.parameter?.acao||'');if(acao==='ping')return json_({ok:true,planilhaId:PLANILHA_ID,abaProdutos:CFG.ABAS.PRODUTOS,totalProdutos:(planilha_().getSheetByName(CFG.ABAS.PRODUTOS)||{getLastRow:()=>1}).getLastRow()-1});return responderPublico_(e)}catch(err){return responder_(e,{ok:false,erro:msg_(err),html:'<div class="sem-produtos">Erro ao carregar os produtos.</div>'})}}
function doPost(e){try{const d=JSON.parse(e.postData.contents||'{}'),a=d.acao;if(a==='login')return json_(login_(d));const u=auth_(d.token);const mapa={verificarToken:()=>({ok:true,usuario:publicUser_(u)}),dashboard:()=>dashboard_(u),listarProdutos:()=>listarProdutos_(u),salvarProduto:()=>salvarProduto_(u,d),alterarStatusProduto:()=>alterarStatusProduto_(u,d),alterarPromocaoProduto:()=>alterarPromocaoProduto_(u,d),excluirProduto:()=>excluirProduto_(u,d),cadastrarVenda:()=>cadastrarVenda_(u,d),listarVendas:()=>listarVendas_(u),listarComissoes:()=>listarComissoes_(u),excluirVenda:()=>excluirVenda_(u,d),listarUsuarios:()=>listarUsuarios_(u),salvarUsuario:()=>salvarUsuario_(u,d)};if(!mapa[a])throw new Error('Ação inválida.');return json_(mapa[a]())}catch(err){return json_({ok:false,erro:msg_(err)})}}
function responderPublico_(e){
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  if(!sh)return responder_(e,{ok:false,html:'<div class="sem-produtos">A aba Produtos não foi encontrada.</div>'});
  const arr=objetos_(sh).map(produtoNormalizado_).filter(x=>x.nome&&['Publicado','Reservado'].includes(x.status)&&x.quantidade>0).sort((a,b)=>Number(ehSim_(b.promocao))-Number(ehSim_(a.promocao)));
  const html=arr.map((x,i)=>{
    const img=x.fotos[0]||'IMG_7025.png',preco=x.preco,promo=ehSim_(x.promocao),reservado=x.status==='Reservado';
    const nome=esc_(x.nome),cat=esc_(x.categoria),cla=esc_(x.classificacao),desc=esc_(x.descricao);
    const mensagem=encodeURIComponent('Olá! Tenho interesse no produto: '+x.nome+(x.classificacao?' - Classificação '+x.classificacao:''));
    const link=x.whatsapp||('https://wa.me/5511960688383?text='+mensagem);
    const acao=reservado?'<button class="whatsapp produto-reservado-btn" type="button" disabled>Produto reservado</button>':`<a class="whatsapp" href="${esc_(link)}" target="_blank" rel="noopener noreferrer">Comprar pelo WhatsApp</a>`;
    return `<article class="produto ${promo?'produto-promocao':''} ${reservado?'produto-reservado':''}" data-status="${reservado?'reservado':'publicado'}" data-categoria="${cat}" data-categoria-nome="${cat}" data-promocao="${promo?'sim':'nao'}" data-oferta="${promo?'sim':'nao'}"><div class="produto-imagem"><img src="${esc_(img)}" alt="${nome}" loading="${i<4?'eager':'lazy'}">${reservado?'<span class="reserved-badge">RESERVADO</span>':''}</div><div class="produto-conteudo"><span class="produto-categoria">${cat}</span><h3>${nome}</h3>${cla?`<span class="produto-classificacao">Classificação ${cla}</span>`:''}${desc?`<p class="produto-descricao">${desc}</p>`:''}<span class="preco-original" data-preco="${preco}">R$ ${preco.toLocaleString('pt-BR',{minimumFractionDigits:2})}</span>${acao}</div></article>`;
  }).join('');
  return responder_(e,{ok:true,html:html||'<div class="sem-produtos">Nenhum produto cadastrado.</div>'});
}
function responder_(e,p){const cb=String(e?.parameter?.callback||'');if(cb&&/^[\w$]+$/.test(cb))return ContentService.createTextOutput(cb+'('+JSON.stringify(p)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return HtmlService.createHtmlOutput(p.html||'').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function login_(d){const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS);if(!sh)throw new Error('Execute configurarSistema() primeiro.');const us=objetos_(sh),u=us.find(x=>String(x['Usuário']).toLowerCase()===String(d.usuario).toLowerCase()&&x['Senha Hash']===hash_(d.senha)&&x.Status==='Ativo');if(!u)throw new Error('Usuário ou senha inválidos.');const token=Utilities.getUuid()+Utilities.getUuid(),val=new Date(Date.now()+CFG.TOKEN_HORAS*3600000);atualizarLinha_(sh,u._linha,{'Token':token,'Validade token':val});return {ok:true,token,usuario:publicUser_(u)}}
function auth_(token){if(!token)throw new Error('Sessão expirada.');const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS),u=objetos_(sh).find(x=>x.Token===token&&x.Status==='Ativo');if(!u||new Date(u['Validade token']).getTime()<Date.now())throw new Error('Sessão expirada.');return u}function admin_(u){if(u.Perfil!=='Administrador')throw new Error('Acesso permitido somente ao administrador.')}function publicUser_(u){return{id:u.ID,nome:u.Nome,usuario:u['Usuário'],perfil:u.Perfil}}
function listarProdutos_(){const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);return{ok:true,produtos:objetos_(sh).map(produtoNormalizado_).filter(x=>x.nome)}}
function salvarProduto_(u,d){
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  const todos=objetos_(sh);
  let r=Number(d.linha)>0?todos.find(x=>Number(x._linha)===Number(d.linha)):null;
  if(!r&&d.id)r=todos.find(x=>String(valorCampo_(x,['ID'])||'')===String(d.id));
  let fotos=r?fotos_(r):[];
  if(Array.isArray(d.fotosUpload)&&d.fotosUpload.length)fotos=salvarFotos_(d.fotosUpload,d.nome);
  const id=r?(valorCampo_(r,['ID'])||uid_('PROD')):uid_('PROD');
  const obj={'Produto':d.nome,'Categoria':d.categoria,'Preço':num_(d.preco),'Preço de custo':num_(d.precoCusto),'Classificação':d.classificacao||'A','Descrição':d.descricao||'','Link da Imagem':fotos[0]||'','Promoção':d.promocao!==undefined&&d.promocao!==''?d.promocao:(r?(valorCampo_(r,['Promoção','Promocao'])||'Não'):'Não'),'WhatsApp':r?(valorCampo_(r,['WhatsApp','Whatsapp'])||''):'','ID':id,'Marca':d.marca||'','Modelo':d.modelo||'','Quantidade':num_(d.quantidade),'Status':d.status||'Publicado','Fotos':JSON.stringify(fotos),'Data de cadastro':r?(valorCampo_(r,['Data de cadastro'])||new Date()):new Date()};
  if(r){
    atualizarLinha_(sh,r._linha,obj);
    const info=cabecalhos_(sh);
    const colPreco=indiceCabecalho_(info.cab,['Preço','Preco']);
    if(colPreco>=0){
      sh.getRange(r._linha,colPreco+1).setValue(num_(d.preco));
    }
  }else appendObj_(sh,obj);
  SpreadsheetApp.flush();
  return{ok:true,id:id,linha:r?r._linha:sh.getLastRow(),preco:num_(d.preco),precoCusto:num_(d.precoCusto)};
}
function alterarStatusProduto_(u,d){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS),todos=objetos_(sh);let r=todos.find(x=>String(x.ID||'')===String(d.id||''));if(!r&&Number(d.linha)>0)r=todos.find(x=>Number(x._linha)===Number(d.linha));if(!r)throw new Error('Produto não encontrado. Atualize a página e tente novamente.');const status=statusNormalizado_(d.status);atualizarLinha_(sh,r._linha,{Status:status});SpreadsheetApp.flush();return{ok:true,status,linha:r._linha}}
function alterarPromocaoProduto_(u,d){
  admin_(u);
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  const todos=objetos_(sh);
  const idInformado=String(d.id||'').trim();
  const linhaInformada=Number(d.linha||0);
  let r=linhaInformada>0?todos.find(x=>Number(x._linha)===linhaInformada):null;
  if(!r&&idInformado)r=todos.find(x=>String(valorCampo_(x,['ID'])||'').trim()===idInformado);
  if(!r)throw new Error('Produto não encontrado na aba Produtos. Atualize a página e tente novamente.');

  const info=cabecalhos_(sh);
  let coluna=indiceCabecalho_(info.cab,['Promoção','Promocao']);
  if(coluna<0){
    coluna=info.cab.length;
    sh.getRange(info.linha,coluna+1).setValue('Promoção');
  }

  const promocao=ehSim_(d.promocao)?'Sim':'Não';
  sh.getRange(r._linha,coluna+1).setValue(promocao);
  SpreadsheetApp.flush();

  const confirmado=sh.getRange(r._linha,coluna+1).getDisplayValue();
  if(ehSim_(confirmado)!==ehSim_(promocao))throw new Error('A promoção não foi gravada na planilha.');
  return{ok:true,promocao:ehSim_(confirmado)?'Sim':'Não',linha:r._linha,id:String(valorCampo_(r,['ID'])||'')};
}
function excluirProduto_(u,d){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS),r=objetos_(sh).find(x=>x.ID===d.id);if(!r)throw new Error('Produto não encontrado.');sh.deleteRow(r._linha);log_(u,'Produto',d.id,d.motivo);return{ok:true}}

function apenasDigitos_(v){return String(v||'').replace(/\D/g,'')}
function cpfValido_(cpf){cpf=apenasDigitos_(cpf);if(cpf.length!==11||/^(\d)\1{10}$/.test(cpf))return false;let soma=0;for(let i=0;i<9;i++)soma+=Number(cpf[i])*(10-i);let d=(soma*10)%11;if(d===10)d=0;if(d!==Number(cpf[9]))return false;soma=0;for(let i=0;i<10;i++)soma+=Number(cpf[i])*(11-i);d=(soma*10)%11;if(d===10)d=0;return d===Number(cpf[10])}
function cnpjValido_(cnpj){cnpj=apenasDigitos_(cnpj);if(cnpj.length!==14||/^(\d)\1{13}$/.test(cnpj))return false;function calc(p){let soma=0;for(let i=0;i<p.length;i++)soma+=Number(cnpj[i])*p[i];const r=soma%11;return r<2?0:11-r}const d1=calc([5,4,3,2,9,8,7,6,5,4,3,2]);if(d1!==Number(cnpj[12]))return false;return calc([6,5,4,3,2,9,8,7,6,5,4,3,2])===Number(cnpj[13])}
function documentoValido_(v){const d=apenasDigitos_(v);return !d||(d.length===11?cpfValido_(d):d.length===14&&cnpjValido_(d))}
function telefoneValido_(v){const d=apenasDigitos_(v);return !d||((d.length===10||d.length===11)&&d[0]!=='0'&&d[1]!=='0')}

function cadastrarVenda_(u,d){
  const ss=planilha_();
  const ps=ss.getSheetByName(CFG.ABAS.PRODUTOS);
  const vs=criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria']);
  if(!ps)throw new Error('A aba Produtos não foi encontrada.');
  const avulsa=String(d.tipoVenda||'estoque')==='avulsa';
  const qtd=Math.max(1,num_(d.quantidade));
  let p=null,produto='',marca='',modelo='',categoria='',idProduto='',custoUnitario=0;
  if(avulsa){
    produto=String(d.produtoManual||'').trim();
    if(!produto)throw new Error('Informe o produto vendido.');
    marca=String(d.marcaManual||'').trim();modelo=String(d.modeloManual||'').trim();categoria=String(d.categoriaManual||'').trim();
    produto=[produto,marca,modelo].filter(Boolean).join(' • ');
  }else{
    const idSelecionado=String(d.idProduto||'').trim();
    const linhaSelecionada=Number(idSelecionado.replace(/^LINHA-/i,''))||Number(d.linhaProduto||0);
    p=objetos_(ps).find(x=>String(valorCampo_(x,['ID'])||'')===idSelecionado || (linhaSelecionada>0&&Number(x._linha)===linhaSelecionada));
    if(!p)throw new Error('Produto selecionado não foi encontrado no estoque. Atualize a página e selecione novamente.');
    const pn=produtoNormalizado_(p);
    if(pn.quantidade<qtd)throw new Error('Quantidade insuficiente no estoque.');
    idProduto=pn.id;produto=pn.nome;marca=pn.marca||'';modelo=pn.modelo||'';categoria=pn.categoria||'';custoUnitario=num_(pn.precoCusto);
  }
  const documento=apenasDigitos_(d.documento||'');if(!documentoValido_(documento))throw new Error('CPF ou CNPJ inválido.');
  const telefone=apenasDigitos_(d.telefone||'');if(!telefoneValido_(telefone))throw new Error('Telefone inválido. Informe DDD e número.');
  const valor=num_(d.valorVenda);if(valor<=0)throw new Error('Informe o valor da venda.');
  const id=uid_('VEN'),data=d.dataVenda||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const custoTotal=Number((custoUnitario*qtd).toFixed(2));
  const lucro=Number((valor-custoTotal).toFixed(2));
  const obj={'ID Venda':id,'Data da venda':data,'Data de cadastro':new Date(),'Cliente':d.cliente,'CPF/CNPJ':documento,'Telefone':telefone,'ID Produto':idProduto,'Produto':produto,'Quantidade':qtd,'Valor da venda':valor,'Custo unitário':custoUnitario,'Custo total':custoTotal,'Lucro':lucro,'Forma de pagamento':d.pagamento,'Garantia dias':num_(d.garantia),'Observações':d.observacoes||'','Vendedora':u.Nome,'Origem do produto':avulsa?'Venda avulsa':'Estoque','Marca':marca,'Modelo':modelo,'Categoria':categoria};
  appendObj_(vs,obj);
  if(p){const pn=produtoNormalizado_(p),nova=pn.quantidade-qtd;atualizarLinha_(ps,p._linha,{Quantidade:nova,Status:nova<=0?'Oculto':pn.status});}
  return{ok:true,venda:{id,cliente:d.cliente,documento,telefone,produto,quantidade:qtd,valor,custoUnitario,custoTotal,lucro,pagamento:d.pagamento,garantia:num_(d.garantia),observacoes:d.observacoes||'',data,vendedora:u.Nome,origem:avulsa?'Venda avulsa':'Estoque'}};
}
function listarVendas_(){
  const sh=planilha_().getSheetByName(CFG.ABAS.VENDAS);
  const custos={};
  listarProdutos_().produtos.forEach(p=>{custos[String(p.id||'')]=num_(p.precoCusto)});
  const vendas=objetos_(sh).reverse().map(x=>{
    const quantidade=num_(x.Quantidade);
    const valor=num_(x['Valor da venda']);
    const idProduto=String(x['ID Produto']||'');
    const custoSalvo=x['Custo total'];
    const custoUnitarioSalvo=x['Custo unitário'];
    const custoUnitario=(custoUnitarioSalvo!==''&&custoUnitarioSalvo!==undefined)?num_(custoUnitarioSalvo):num_(custos[idProduto]);
    const custoTotal=(custoSalvo!==''&&custoSalvo!==undefined)?num_(custoSalvo):Number((custoUnitario*quantidade).toFixed(2));
    const lucro=(x['Lucro']!==''&&x['Lucro']!==undefined)?num_(x['Lucro']):Number((valor-custoTotal).toFixed(2));
    return{id:x['ID Venda'],data:fmtData_(x['Data da venda']),cliente:x.Cliente,documento:x['CPF/CNPJ'],telefone:x.Telefone,idProduto,produto:x.Produto,quantidade,valor,custoUnitario,custoTotal,lucro,pagamento:x['Forma de pagamento'],garantia:num_(x['Garantia dias']),observacoes:x['Observações'],vendedora:x.Vendedora,origem:x['Origem do produto']||((x['ID Produto'])?'Estoque':'Venda avulsa')};
  });
  return{ok:true,vendas};
}

function listarComissoes_(u){
  const percentual=0.02;
  const hoje=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const mes=hoje.slice(0,7);
  let vendas=listarVendas_().vendas;
  if(u.Perfil!=='Administrador')vendas=vendas.filter(v=>String(v.vendedora||'').trim().toLowerCase()===String(u.Nome||'').trim().toLowerCase());
  const itens=vendas.map(v=>({...v,comissao:Number((num_(v.valor)*percentual).toFixed(2))}));
  const vendasHoje=itens.filter(v=>isoData_(v.data)===hoje);
  const vendasMes=itens.filter(v=>isoData_(v.data).slice(0,7)===mes);
  return{ok:true,percentual:2,totalHoje:vendasHoje.reduce((s,v)=>s+v.comissao,0),totalMes:vendasMes.reduce((s,v)=>s+v.comissao,0),totalGeral:itens.reduce((s,v)=>s+v.comissao,0),vendasMes:vendasMes.length,itens};
}

function excluirVenda_(u,d){admin_(u);const ss=planilha_(),vs=ss.getSheetByName(CFG.ABAS.VENDAS),r=objetos_(vs).find(x=>x['ID Venda']===d.id);if(!r)throw new Error('Venda não encontrada.');const ps=ss.getSheetByName(CFG.ABAS.PRODUTOS),p=objetos_(ps).find(x=>String(x.ID)===String(r['ID Produto']));if(p)atualizarLinha_(ps,p._linha,{Quantidade:produtoNormalizado_(p).quantidade+num_(r.Quantidade)});vs.deleteRow(r._linha);log_(u,'Venda',d.id,d.motivo);return{ok:true}}
function dashboard_(){const hoje=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'),mes=hoje.slice(0,7),v=listarVendas_().vendas;const vh=v.filter(x=>isoData_(x.data)===hoje),vm=v.filter(x=>isoData_(x.data).slice(0,7)===mes);const custoHoje=vh.reduce((s,x)=>s+num_(x.custoTotal),0),custoMes=vm.reduce((s,x)=>s+num_(x.custoTotal),0),faturamentoHoje=vh.reduce((s,x)=>s+x.valor,0),faturamentoMes=vm.reduce((s,x)=>s+x.valor,0);return{ok:true,faturamentoHoje,faturamentoMes,custoHoje,custoMes,lucroHoje:faturamentoHoje-custoHoje,lucroMes:faturamentoMes-custoMes,vendasMes:vm.length,estoque:listarProdutos_().produtos.reduce((s,x)=>s+x.quantidade,0),ultimasVendas:v.slice(0,8)}}
function listarUsuarios_(u){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS);return{ok:true,usuarios:objetos_(sh).map(x=>({id:x.ID,nome:x.Nome,usuario:x['Usuário'],perfil:x.Perfil,status:x.Status}))}}
function salvarUsuario_(u,d){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS),todos=objetos_(sh),r=d.id?todos.find(x=>x.ID===d.id):null;if(todos.some(x=>String(x['Usuário']).toLowerCase()===String(d.usuario).toLowerCase()&&x.ID!==d.id))throw new Error('Esse usuário já existe.');const obj={'ID':r?r.ID:uid_('USR'),'Nome':d.nome,'Usuário':d.usuario,'Senha Hash':d.senha?hash_(d.senha):(r?r['Senha Hash']:''),'Perfil':d.perfil||'Vendedora','Status':d.status||'Ativo','Token':r?r.Token:'','Validade token':r?r['Validade token']:'','Data de cadastro':r?r['Data de cadastro']:new Date()};if(r)atualizarLinha_(sh,r._linha,obj);else appendObj_(sh,obj);return{ok:true}}
function normalizarChave_(v){return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase().replace(/\s+/g,' ')}
function indiceCabecalho_(cab,nomes){const alvos=nomes.map(normalizarChave_);return cab.findIndex(x=>alvos.includes(normalizarChave_(x)))}
function valorCampo_(obj,nomes){const alvos=nomes.map(normalizarChave_);for(const k of Object.keys(obj||{})){if(alvos.includes(normalizarChave_(k)))return obj[k]}return''}
function produtoNormalizado_(x){
  const nome=String(valorCampo_(x,['Produto','Nome'])||'').trim();
  return{
    id:String(valorCampo_(x,['ID'])||('LINHA-'+Number(x._linha||0))),
    linha:Number(x._linha||0),
    nome,
    categoria:String(valorCampo_(x,['Categoria'])||'outros'),
    preco:num_(valorCampo_(x,['Preço','Preco'])),
    precoCusto:num_(valorCampo_(x,['Preço de custo','Preco de custo','Custo'])),
    classificacao:String(valorCampo_(x,['Classificação','Classificacao'])||''),
    descricao:String(valorCampo_(x,['Descrição','Descricao'])||''),
    marca:String(valorCampo_(x,['Marca'])||''),
    modelo:String(valorCampo_(x,['Modelo'])||''),
    quantidade:valorCampo_(x,['Quantidade'])===''?1:Math.max(0,num_(valorCampo_(x,['Quantidade']))),
    status:statusNormalizado_(valorCampo_(x,['Status'])||'Publicado'),
    promocao:valorCampo_(x,['Promoção','Promocao']),
    whatsapp:String(valorCampo_(x,['WhatsApp','Whatsapp','Link WhatsApp'])||''),
    fotos:fotos_(x),
    _linha:x._linha
  };
}
function migrarProdutosExistentes_(sh){
  if(!sh||sh.getLastRow()<2)return;

  const info=cabecalhos_(sh);
  const range=sh.getRange(info.linha,1,sh.getLastRow()-info.linha+1,sh.getLastColumn());
  const dados=range.getValues();
  const cab=dados[0].map(v=>String(v).trim());
  const col={};
  cab.forEach((nome,i)=>col[nome]=i);

  const agora=new Date();
  let alterou=false;

  for(let i=1;i<dados.length;i++){
    const linha=dados[i];
    const nome=String(linha[col.Nome]||linha[col.Produto]||'').trim();
    if(!nome)continue;

    if(col.ID!==undefined&&!linha[col.ID]){
      linha[col.ID]=uid_('PROD');
      alterou=true;
    }
    if(col.Quantidade!==undefined&&(linha[col.Quantidade]===''||linha[col.Quantidade]===null)){
      linha[col.Quantidade]=1;
      alterou=true;
    }
    if(col.Status!==undefined&&!linha[col.Status]){
      linha[col.Status]='Publicado';
      alterou=true;
    }
    const colImagem=col['Link da Imagem']!==undefined?col['Link da Imagem']:col.Imagem;
    if(col.Fotos!==undefined&&!linha[col.Fotos]&&colImagem!==undefined&&linha[colImagem]){
      linha[col.Fotos]=JSON.stringify([normalizarImagem_(linha[colImagem])]);
      alterou=true;
    }
    if(col['Data de cadastro']!==undefined&&!linha[col['Data de cadastro']]){
      linha[col['Data de cadastro']]=agora;
      alterou=true;
    }
  }

  // Faz somente uma gravação na planilha, mesmo que existam centenas de produtos.
  if(alterou)range.setValues(dados);
}
function salvarFotos_(lista,nome){const pasta=getPasta_();return lista.slice(0,6).map((x,i)=>{const m=String(x.data||'').match(/^data:([^;]+);base64,(.+)$/);if(!m)return'';const arq=pasta.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],(x.nome||nome+'-'+(i+1))));arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return'https://drive.google.com/thumbnail?id='+arq.getId()+'&sz=w1600'}).filter(Boolean)}function getPasta_(){const it=DriveApp.getFoldersByName(CFG.PASTA);return it.hasNext()?it.next():DriveApp.createFolder(CFG.PASTA)}
function objetos_(sh){
  if(!sh)return[];
  const info=cabecalhos_(sh),last=sh.getLastRow();
  if(last<=info.linha)return[];
  const h=info.cab.map(v=>String(v).trim());
  const v=sh.getRange(info.linha+1,1,last-info.linha,h.length).getValues();
  return v.map((r,i)=>{
    const o={_linha:info.linha+1+i};
    h.forEach((x,j)=>{
      if(!x)return;
      // Quando existem colunas duplicadas, mantém o primeiro valor preenchido.
      // Isso evita que uma coluna antiga vazia apague Preço/Promoção na leitura.
      if(!(x in o)||String(o[x]??'').trim()==='')o[x]=r[j];
    });
    return o;
  }).filter(o=>Object.keys(o).some(k=>k!=='_linha'&&String(o[k]??'').trim()!==''));
}
function atualizarLinha_(sh,linha,obj){
  const info=cabecalhos_(sh),h=info.cab.map(v=>String(v).trim());
  const row=sh.getRange(linha,1,1,h.length).getValues()[0];
  Object.keys(obj).forEach(k=>{
    const alvo=normalizarChave_(k);
    const c=h.findIndex(x=>normalizarChave_(x)===alvo);
    if(c>=0)row[c]=obj[k];
  });
  sh.getRange(linha,1,1,h.length).setValues([row]);
}
function appendObj_(sh,obj){
  if(!sh)throw new Error('A aba necessária não foi encontrada.');
  const info=cabecalhos_(sh);
  const h=info.cab.map(v=>String(v).trim());
  let ultima=h.length;
  while(ultima>0&&!h[ultima-1])ultima--;
  if(!ultima)throw new Error('A aba está sem cabeçalho.');
  const linha=Math.max(sh.getLastRow()+1,info.linha+1);
  const valores=h.slice(0,ultima).map(k=>Object.prototype.hasOwnProperty.call(obj,k)?obj[k]:'');
  sh.getRange(linha,1,1,ultima).setValues([valores]);
  SpreadsheetApp.flush();
}
function fotos_(x){try{const a=JSON.parse(x.Fotos||'[]');if(Array.isArray(a)&&a.length)return a.filter(Boolean)}catch(e){}const img=x['Link da Imagem']||x.Imagem;return img?[normalizarImagem_(img)]:[]}function normalizarImagem_(url){const s=String(url||'').trim(),m=s.match(/[-\w]{25,}/);return s.includes('drive.google.com')&&m?'https://drive.google.com/thumbnail?id='+m[0]+'&sz=w1600':s}function num_(v){let s=String(v??'').replace(/R\$/g,'').trim();if(s.includes(','))s=s.replace(/\./g,'').replace(',','.');return Number(s)||0}function statusNormalizado_(v){const s=String(v||'Publicado').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();if(s==='reservado')return'Reservado';if(s==='oculto')return'Oculto';return'Publicado'}function ehSim_(v){return['sim','true','1','yes','x'].includes(String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase())}function hash_(s){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s)))}function uid_(p){return p+'-'+Utilities.getUuid().slice(0,8).toUpperCase()}function msg_(e){return e&&e.message?e.message:String(e)}function esc_(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;')}function log_(u,t,id,m){appendObj_(planilha_().getSheetByName(CFG.ABAS.LOG),{'Data':new Date(),'Tipo':t,'ID':id,'Motivo':m,'Usuário':u.Nome})}function fmtData_(d){if(d instanceof Date)return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy');const s=String(d||'');if(/^\d{4}-\d{2}-\d{2}$/.test(s)){const[a,b,c]=s.split('-');return`${c}/${b}/${a}`}return s}function isoData_(s){const v=String(s||'');if(/^\d{2}\/\d{2}\/\d{4}$/.test(v)){const[d,m,a]=v.split('/');return`${a}-${m}-${d}`}return v.slice(0,10)}

/* ===== CORREÇÃO DEFINITIVA DA ESTRUTURA DA ABA PRODUTOS ===== */
const CABECALHO_PRODUTOS_FINAL_ = [
  'Produto','Categoria','Preço','Preço de custo','Classificação','Descrição','Link da Imagem','Promoção',
  'WhatsApp','ID','Marca','Modelo','Quantidade','Status','Fotos','Data de cadastro'
];

function configurarSistema(){
  const ss=planilha_();
  let produtos=ss.getSheetByName(CFG.ABAS.PRODUTOS);
  if(!produtos){
    produtos=ss.insertSheet(CFG.ABAS.PRODUTOS);
    produtos.getRange(1,1,1,CABECALHO_PRODUTOS_FINAL_.length).setValues([CABECALHO_PRODUTOS_FINAL_]);
  }
  corrigirEstruturaProdutos_();
  criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria']);
  criarAba_(ss,CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro']);
  criarAba_(ss,CFG.ABAS.LOG,['Data','Tipo','ID','Motivo','Usuário']);
  migrarProdutosExistentes_(produtos);
  const sh=ss.getSheetByName(CFG.ABAS.USUARIOS);
  if(sh.getLastRow()<2)sh.appendRow([uid_('USR'),'Administrador','admin',hash_('123456'),'Administrador','Ativo','','',new Date()]);
  return 'Sistema configurado e aba Produtos corrigida.';
}

function criarAba_(ss,nome,cab){
  let sh=ss.getSheetByName(nome);
  if(!sh){
    sh=ss.insertSheet(nome);
    sh.getRange(1,1,1,cab.length).setValues([cab]);
    sh.setFrozenRows(1);
    return sh;
  }
  if(nome===CFG.ABAS.PRODUTOS)return sh;
  const info=cabecalhos_(sh);
  const atual=info.cab.map(v=>String(v).trim());
  cab.forEach(c=>{
    if(indiceCabecalho_(atual,[c])<0){
      sh.getRange(info.linha,atual.length+1).setValue(c);
      atual.push(c);
    }
  });
  sh.setFrozenRows(info.linha);
  return sh;
}

function corrigirEstruturaProdutos_(){
  const ss=planilha_();
  const sh=ss.getSheetByName(CFG.ABAS.PRODUTOS);
  if(!sh)throw new Error('A aba Produtos não foi encontrada.');

  const info=cabecalhos_(sh);
  const lastRow=Math.max(sh.getLastRow(),info.linha);
  const lastCol=Math.max(sh.getLastColumn(),CABECALHO_PRODUTOS_FINAL_.length);
  const dados=sh.getRange(info.linha,1,lastRow-info.linha+1,lastCol).getValues();
  const cab=dados[0].map(v=>String(v).trim());

  // Cria uma cópia de segurança somente na primeira correção.
  const backupNome='Produtos_backup_antes_correcao';
  if(!ss.getSheetByName(backupNome)){
    const backup=sh.copyTo(ss).setName(backupNome);
    backup.hideSheet();
  }

  const indices=(nomes)=>{
    const alvos=nomes.map(normalizarChave_);
    const arr=[];
    cab.forEach((v,i)=>{if(alvos.includes(normalizarChave_(v)))arr.push(i)});
    return arr;
  };
  const primeiro=(linha,nomes)=>{
    for(const i of indices(nomes)){
      const v=linha[i];
      if(v!==''&&v!==null&&v!==undefined)return v;
    }
    return '';
  };
  const ultimo=(linha,nomes)=>{
    const arr=indices(nomes).slice().reverse();
    for(const i of arr){
      const v=linha[i];
      if(v!==''&&v!==null&&v!==undefined)return v;
    }
    return '';
  };
  const promocao=(linha)=>{
    const vals=indices(['Promoção','Promocao']).map(i=>linha[i]);
    return vals.some(ehSim_)?'Sim':'Não';
  };

  const saida=[CABECALHO_PRODUTOS_FINAL_];
  for(let r=1;r<dados.length;r++){
    const linha=dados[r];
    const produto=primeiro(linha,['Produto','Nome']);
    if(String(produto||'').trim()==='')continue;
    const linkImagem=primeiro(linha,['Link da Imagem','Imagem']);
    let fotos=primeiro(linha,['Fotos']);
    if(!fotos&&linkImagem)fotos=JSON.stringify([normalizarImagem_(linkImagem)]);
    saida.push([
      produto,
      primeiro(linha,['Categoria'])||'outros',
      num_(ultimo(linha,['Preço','Preco'])),
      num_(ultimo(linha,['Preço de custo','Preco de custo','Custo'])),
      primeiro(linha,['Classificação','Classificacao'])||'A',
      primeiro(linha,['Descrição','Descricao']),
      linkImagem,
      promocao(linha),
      primeiro(linha,['WhatsApp','Whatsapp','Link WhatsApp']),
      primeiro(linha,['ID'])||uid_('PROD'),
      primeiro(linha,['Marca']),
      primeiro(linha,['Modelo']),
      primeiro(linha,['Quantidade'])===''?1:Math.max(0,num_(primeiro(linha,['Quantidade']))),
      statusNormalizado_(primeiro(linha,['Status'])||'Publicado'),
      fotos||'[]',
      primeiro(linha,['Data de cadastro'])||new Date()
    ]);
  }

  // Remove validações antigas antes de reorganizar, para não bloquear a gravação.
  const area=sh.getRange(info.linha,1,Math.max(lastRow-info.linha+1,saida.length),lastCol);
  area.clearDataValidations();
  area.clearContent();

  // Regrava uma única estrutura oficial.
  sh.getRange(info.linha,1,saida.length,CABECALHO_PRODUTOS_FINAL_.length).setValues(saida);
  // A coluna Preço pertence a uma tabela tipada do Google Sheets.
  // O formato é controlado pela própria tabela; não aplicar setNumberFormat aqui.

  // Recria a validação correta da coluna Promoção.
  if(saida.length>1){
    const regraPromocao=SpreadsheetApp.newDataValidation()
      .requireValueInList(['Sim','Não'],true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(info.linha+1,8,saida.length-1,1).setDataValidation(regraPromocao);
  }
  sh.setFrozenRows(info.linha);
  SpreadsheetApp.flush();
  return 'Estrutura da aba Produtos corrigida com '+(saida.length-1)+' produtos.';
}

function alterarPromocaoProduto_(u,d){
  admin_(u);
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  const todos=objetos_(sh);
  const linha=Number(d.linha||0);
  const id=String(d.id||'').trim();
  let r=linha>0?todos.find(x=>Number(x._linha)===linha):null;
  if(!r&&id)r=todos.find(x=>String(x.ID||'').trim()===id);
  if(!r)throw new Error('Produto não encontrado.');
  const info=cabecalhos_(sh);
  const col=indiceCabecalho_(info.cab,['Promoção','Promocao']);
  if(col<0)throw new Error('Coluna Promoção não encontrada. Execute configurarSistema().');
  const valor=ehSim_(d.promocao)?'Sim':'Não';
  sh.getRange(r._linha,col+1).setValue(valor);
  SpreadsheetApp.flush();
  return {ok:true,promocao:valor,linha:r._linha,id:String(r.ID||'')};
}
