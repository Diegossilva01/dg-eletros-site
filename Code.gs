const CFG={ABAS:{PRODUTOS:'Produtos',VENDAS:'Vendas',USUARIOS:'Usuarios',LOG:'LogExclusoes'},TOKEN_HORAS:12,PASTA:'DG Eletros - Fotos dos Produtos'};
const PLANILHA_ID='1FCiKWKy_DMa3Dwj12cCD2KM_qYFuTnRlKSQ9fM2FleA';
function planilha_(){return SpreadsheetApp.openById(PLANILHA_ID);}
function configurarSistema(){const ss=planilha_();criarAba_(ss,CFG.ABAS.PRODUTOS,['Nome','Categoria','Preço','Classificação','Descrição','Imagem','WhatsApp','ID','Promoção','Marca','Modelo','Quantidade','Status','Fotos','Data de cadastro']);criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria']);criarAba_(ss,CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro']);criarAba_(ss,CFG.ABAS.LOG,['Data','Tipo','ID','Motivo','Usuário']);migrarProdutosExistentes_(ss.getSheetByName(CFG.ABAS.PRODUTOS));const sh=ss.getSheetByName(CFG.ABAS.USUARIOS);if(sh.getLastRow()<2)sh.appendRow([uid_('USR'),'Administrador','admin',hash_('123456'),'Administrador','Ativo','','',new Date()]);return 'Sistema configurado. Login inicial: admin / 123456';}
function criarAba_(ss,nome,cab){let sh=ss.getSheetByName(nome);if(!sh)sh=ss.insertSheet(nome);if(sh.getLastRow()===0)sh.appendRow(cab);else{const atual=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];cab.forEach(c=>{if(!atual.includes(c)){sh.getRange(1,sh.getLastColumn()+1).setValue(c);atual.push(c)}})}sh.setFrozenRows(1);return sh}
function doGet(e){try{const acao=String(e?.parameter?.acao||'');if(acao==='ping')return json_({ok:true});return responderPublico_(e)}catch(err){return responder_(e,{ok:false,erro:msg_(err),html:'<div class="sem-produtos">Erro ao carregar os produtos.</div>'})}}
function doPost(e){try{const d=JSON.parse(e.postData.contents||'{}'),a=d.acao;if(a==='login')return json_(login_(d));const u=auth_(d.token);const mapa={verificarToken:()=>({ok:true,usuario:publicUser_(u)}),dashboard:()=>dashboard_(u),listarProdutos:()=>listarProdutos_(u),salvarProduto:()=>salvarProduto_(u,d),alterarStatusProduto:()=>alterarStatusProduto_(u,d),excluirProduto:()=>excluirProduto_(u,d),cadastrarVenda:()=>cadastrarVenda_(u,d),listarVendas:()=>listarVendas_(u),excluirVenda:()=>excluirVenda_(u,d),listarUsuarios:()=>listarUsuarios_(u),salvarUsuario:()=>salvarUsuario_(u,d)};if(!mapa[a])throw new Error('Ação inválida.');return json_(mapa[a]())}catch(err){return json_({ok:false,erro:msg_(err)})}}
function responderPublico_(e){
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  if(!sh)return responder_(e,{ok:false,html:'<div class="sem-produtos">A aba Produtos não foi encontrada.</div>'});
  const arr=objetos_(sh).map(produtoNormalizado_).filter(x=>x.nome&&x.status==='Publicado'&&x.quantidade>0);
  const html=arr.map((x,i)=>{
    const img=x.fotos[0]||'IMG_7025.png',preco=x.preco,promo=ehSim_(x.promocao);
    const nome=esc_(x.nome),cat=esc_(x.categoria),cla=esc_(x.classificacao),desc=esc_(x.descricao);
    const mensagem=encodeURIComponent('Olá! Tenho interesse no produto: '+x.nome+(x.classificacao?' - Classificação '+x.classificacao:''));
    const link=x.whatsapp||('https://wa.me/5511960688383?text='+mensagem);
    return `<article class="produto ${promo?'produto-promocao':''}" data-categoria="${cat}" data-categoria-nome="${cat}" data-promocao="${promo?'sim':'nao'}" data-oferta="${promo?'sim':'nao'}"><div class="produto-imagem"><img src="${esc_(img)}" alt="${nome}" loading="${i<4?'eager':'lazy'}"></div><div class="produto-conteudo"><span class="produto-categoria">${cat}</span><h3>${nome}</h3>${cla?`<span class="produto-classificacao">Classificação ${cla}</span>`:''}${desc?`<p class="produto-descricao">${desc}</p>`:''}<span class="preco-original" data-preco="${preco}">R$ ${preco.toLocaleString('pt-BR',{minimumFractionDigits:2})}</span><a class="whatsapp" href="${esc_(link)}" target="_blank" rel="noopener noreferrer">Comprar pelo WhatsApp</a></div></article>`;
  }).join('');
  return responder_(e,{ok:true,html:html||'<div class="sem-produtos">Nenhum produto cadastrado.</div>'});
}
function responder_(e,p){const cb=String(e?.parameter?.callback||'');if(cb&&/^[\w$]+$/.test(cb))return ContentService.createTextOutput(cb+'('+JSON.stringify(p)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return HtmlService.createHtmlOutput(p.html||'').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function login_(d){const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS);if(!sh)throw new Error('Execute configurarSistema() primeiro.');const us=objetos_(sh),u=us.find(x=>String(x['Usuário']).toLowerCase()===String(d.usuario).toLowerCase()&&x['Senha Hash']===hash_(d.senha)&&x.Status==='Ativo');if(!u)throw new Error('Usuário ou senha inválidos.');const token=Utilities.getUuid()+Utilities.getUuid(),val=new Date(Date.now()+CFG.TOKEN_HORAS*3600000);atualizarLinha_(sh,u._linha,{'Token':token,'Validade token':val});return {ok:true,token,usuario:publicUser_(u)}}
function auth_(token){if(!token)throw new Error('Sessão expirada.');const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS),u=objetos_(sh).find(x=>x.Token===token&&x.Status==='Ativo');if(!u||new Date(u['Validade token']).getTime()<Date.now())throw new Error('Sessão expirada.');return u}function admin_(u){if(u.Perfil!=='Administrador')throw new Error('Acesso permitido somente ao administrador.')}function publicUser_(u){return{id:u.ID,nome:u.Nome,usuario:u['Usuário'],perfil:u.Perfil}}
function listarProdutos_(){const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);return{ok:true,produtos:objetos_(sh).map(produtoNormalizado_).filter(x=>x.nome)}}
function salvarProduto_(u,d){
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  let r=d.id?objetos_(sh).find(x=>String(x.ID)===String(d.id)):null;
  let fotos=r?fotos_(r):[];
  if(Array.isArray(d.fotosUpload)&&d.fotosUpload.length)fotos=salvarFotos_(d.fotosUpload,d.nome);
  const id=r?(r.ID||uid_('PROD')):uid_('PROD');
  const obj={'Nome':d.nome,'Produto':d.nome,'Categoria':d.categoria,'Preço':num_(d.preco),'Classificação':d.classificacao||'A','Descrição':d.descricao||'','Imagem':fotos[0]||'','WhatsApp':r?(r.WhatsApp||r.Whatsapp||''):'','ID':id,'Promoção':d.promocao||'NÃO','Marca':d.marca||'','Modelo':d.modelo||'','Quantidade':num_(d.quantidade),'Status':d.status||'Publicado','Fotos':JSON.stringify(fotos),'Data de cadastro':r?(r['Data de cadastro']||new Date()):new Date()};
  if(r)atualizarLinha_(sh,r._linha,obj);else appendObj_(sh,obj);
  return{ok:true,id:id};
}
function alterarStatusProduto_(u,d){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS),r=objetos_(sh).find(x=>x.ID===d.id);if(!r)throw new Error('Produto não encontrado.');atualizarLinha_(sh,r._linha,{Status:d.status});return{ok:true}}
function excluirProduto_(u,d){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS),r=objetos_(sh).find(x=>x.ID===d.id);if(!r)throw new Error('Produto não encontrado.');sh.deleteRow(r._linha);log_(u,'Produto',d.id,d.motivo);return{ok:true}}
function cadastrarVenda_(u,d){
  const ss=planilha_(),ps=ss.getSheetByName(CFG.ABAS.PRODUTOS),vs=ss.getSheetByName(CFG.ABAS.VENDAS);
  const avulsa=String(d.tipoVenda||'estoque')==='avulsa';
  const qtd=Math.max(1,num_(d.quantidade));
  let p=null,produto='',marca='',modelo='',categoria='',idProduto='';
  if(avulsa){
    produto=String(d.produtoManual||'').trim();
    if(!produto)throw new Error('Informe o produto vendido.');
    marca=String(d.marcaManual||'').trim();modelo=String(d.modeloManual||'').trim();categoria=String(d.categoriaManual||'').trim();
    produto=[produto,marca,modelo].filter(Boolean).join(' • ');
  }else{
    p=objetos_(ps).find(x=>String(x.ID)===String(d.idProduto));
    if(!p)throw new Error('Selecione um produto do estoque.');
    const pn=produtoNormalizado_(p);
    if(pn.quantidade<qtd)throw new Error('Quantidade insuficiente no estoque.');
    idProduto=pn.id;produto=pn.nome;marca=pn.marca||'';modelo=pn.modelo||'';categoria=pn.categoria||'';
  }
  const valor=num_(d.valorVenda);if(valor<=0)throw new Error('Informe o valor da venda.');
  const id=uid_('VEN'),data=d.dataVenda||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const obj={'ID Venda':id,'Data da venda':data,'Data de cadastro':new Date(),'Cliente':d.cliente,'CPF/CNPJ':d.documento||'','Telefone':d.telefone||'','ID Produto':idProduto,'Produto':produto,'Quantidade':qtd,'Valor da venda':valor,'Forma de pagamento':d.pagamento,'Garantia dias':num_(d.garantia),'Observações':d.observacoes||'','Vendedora':u.Nome,'Origem do produto':avulsa?'Venda avulsa':'Estoque','Marca':marca,'Modelo':modelo,'Categoria':categoria};
  appendObj_(vs,obj);
  if(p){const pn=produtoNormalizado_(p),nova=pn.quantidade-qtd;atualizarLinha_(ps,p._linha,{Quantidade:nova,Status:nova<=0?'Oculto':pn.status});}
  return{ok:true,venda:{id,cliente:d.cliente,documento:d.documento||'',telefone:d.telefone||'',produto,quantidade:qtd,valor,pagamento:d.pagamento,garantia:num_(d.garantia),observacoes:d.observacoes||'',data,vendedora:u.Nome,origem:avulsa?'Venda avulsa':'Estoque'}};
}
function listarVendas_(){const sh=planilha_().getSheetByName(CFG.ABAS.VENDAS);return{ok:true,vendas:objetos_(sh).reverse().map(x=>({id:x['ID Venda'],data:fmtData_(x['Data da venda']),cliente:x.Cliente,documento:x['CPF/CNPJ'],telefone:x.Telefone,idProduto:x['ID Produto'],produto:x.Produto,quantidade:num_(x.Quantidade),valor:num_(x['Valor da venda']),pagamento:x['Forma de pagamento'],garantia:num_(x['Garantia dias']),observacoes:x['Observações'],vendedora:x.Vendedora,origem:x['Origem do produto']||((x['ID Produto'])?'Estoque':'Venda avulsa')}))}}
function excluirVenda_(u,d){admin_(u);const ss=planilha_(),vs=ss.getSheetByName(CFG.ABAS.VENDAS),r=objetos_(vs).find(x=>x['ID Venda']===d.id);if(!r)throw new Error('Venda não encontrada.');const ps=ss.getSheetByName(CFG.ABAS.PRODUTOS),p=objetos_(ps).find(x=>String(x.ID)===String(r['ID Produto']));if(p)atualizarLinha_(ps,p._linha,{Quantidade:produtoNormalizado_(p).quantidade+num_(r.Quantidade)});vs.deleteRow(r._linha);log_(u,'Venda',d.id,d.motivo);return{ok:true}}
function dashboard_(){const hoje=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'),mes=hoje.slice(0,7),v=listarVendas_().vendas;const vh=v.filter(x=>isoData_(x.data)===hoje),vm=v.filter(x=>isoData_(x.data).slice(0,7)===mes);return{ok:true,faturamentoHoje:vh.reduce((s,x)=>s+x.valor,0),faturamentoMes:vm.reduce((s,x)=>s+x.valor,0),vendasMes:vm.length,estoque:listarProdutos_().produtos.reduce((s,x)=>s+x.quantidade,0),ultimasVendas:v.slice(0,8)}}
function listarUsuarios_(u){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS);return{ok:true,usuarios:objetos_(sh).map(x=>({id:x.ID,nome:x.Nome,usuario:x['Usuário'],perfil:x.Perfil,status:x.Status}))}}
function salvarUsuario_(u,d){admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS),todos=objetos_(sh),r=d.id?todos.find(x=>x.ID===d.id):null;if(todos.some(x=>String(x['Usuário']).toLowerCase()===String(d.usuario).toLowerCase()&&x.ID!==d.id))throw new Error('Esse usuário já existe.');const obj={'ID':r?r.ID:uid_('USR'),'Nome':d.nome,'Usuário':d.usuario,'Senha Hash':d.senha?hash_(d.senha):(r?r['Senha Hash']:''),'Perfil':d.perfil||'Vendedora','Status':d.status||'Ativo','Token':r?r.Token:'','Validade token':r?r['Validade token']:'','Data de cadastro':r?r['Data de cadastro']:new Date()};if(r)atualizarLinha_(sh,r._linha,obj);else appendObj_(sh,obj);return{ok:true}}
function produtoNormalizado_(x){
  const nome=String(x.Produto||x.Nome||'').trim();
  return{id:String(x.ID||''),nome:nome,categoria:String(x.Categoria||'outros'),preco:num_(x['Preço']||x.Preco),classificacao:String(x['Classificação']||x.Classificacao||''),descricao:String(x['Descrição']||x.Descricao||''),marca:String(x.Marca||''),modelo:String(x.Modelo||''),quantidade:x.Quantidade===''||x.Quantidade===undefined?1:Math.max(0,num_(x.Quantidade)),status:String(x.Status||'Publicado'),promocao:x['Promoção']||x.Promocao||'',whatsapp:String(x.WhatsApp||x.Whatsapp||x['Link WhatsApp']||''),fotos:fotos_(x),_linha:x._linha};
}
function migrarProdutosExistentes_(sh){
  if(!sh||sh.getLastRow()<2)return;

  const range=sh.getDataRange();
  const dados=range.getValues();
  const cab=dados[0];
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
    if(col.Fotos!==undefined&&!linha[col.Fotos]&&col.Imagem!==undefined&&linha[col.Imagem]){
      linha[col.Fotos]=JSON.stringify([normalizarImagem_(linha[col.Imagem])]);
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
function objetos_(sh){if(!sh||sh.getLastRow()<2)return[];const v=sh.getDataRange().getValues(),h=v.shift();return v.map((r,i)=>{const o={_linha:i+2};h.forEach((x,j)=>o[x]=r[j]);return o})}function atualizarLinha_(sh,linha,obj){const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];Object.keys(obj).forEach(k=>{const c=h.indexOf(k);if(c>=0)sh.getRange(linha,c+1).setValue(obj[k])})}function appendObj_(sh,obj){const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];sh.appendRow(h.map(k=>obj[k]??''))}
function fotos_(x){try{const a=JSON.parse(x.Fotos||'[]');if(Array.isArray(a)&&a.length)return a.filter(Boolean)}catch(e){}return x.Imagem?[normalizarImagem_(x.Imagem)]:[]}function normalizarImagem_(url){const s=String(url||'').trim(),m=s.match(/[-\w]{25,}/);return s.includes('drive.google.com')&&m?'https://drive.google.com/thumbnail?id='+m[0]+'&sz=w1600':s}function num_(v){let s=String(v??'').replace(/R\$/g,'').trim();if(s.includes(','))s=s.replace(/\./g,'').replace(',','.');return Number(s)||0}function ehSim_(v){return['sim','true','1','yes','x'].includes(String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase())}function hash_(s){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s)))}function uid_(p){return p+'-'+Utilities.getUuid().slice(0,8).toUpperCase()}function msg_(e){return e&&e.message?e.message:String(e)}function esc_(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;')}function log_(u,t,id,m){appendObj_(planilha_().getSheetByName(CFG.ABAS.LOG),{'Data':new Date(),'Tipo':t,'ID':id,'Motivo':m,'Usuário':u.Nome})}function fmtData_(d){if(d instanceof Date)return Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy');const s=String(d||'');if(/^\d{4}-\d{2}-\d{2}$/.test(s)){const[a,b,c]=s.split('-');return`${c}/${b}/${a}`}return s}function isoData_(s){const v=String(s||'');if(/^\d{2}\/\d{2}\/\d{4}$/.test(v)){const[d,m,a]=v.split('/');return`${a}-${m}-${d}`}return v.slice(0,10)}
