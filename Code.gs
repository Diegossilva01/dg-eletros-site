const CFG={ABAS:{PRODUTOS:'Produtos',VENDAS:'Vendas',USUARIOS:'Usuarios',LOG:'LogExclusoes',COMISSOES:'Pagamentos Comissoes',HORAS_EXTRAS:'Horas Extras'},TOKEN_HORAS:720,PASTA:'DG Eletros - Fotos dos Produtos'};
const PLANILHA_ID='1FCiKWKy_DMa3Dwj12cCD2KM_qYFuTnRlKSQ9fM2FleA';
function planilha_(){return SpreadsheetApp.openById(PLANILHA_ID);}
function configurarSistema(){const ss=planilha_();criarAba_(ss,CFG.ABAS.PRODUTOS,['Produto','Categoria','Preço','Preço de custo','Classificação','Descrição','Link da Imagem','Promoção','WhatsApp','ID','Marca','Modelo','Quantidade','Status','Fotos','Data de cadastro']);criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria','Entrada','Saldo a receber','Parcelas JSON','Status recebimento','Recebimentos JSON','Status comissão','Comissão revisada por','Data revisão comissão']);criarAba_(ss,CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro','Salário mensal','Horas semanais']);criarAba_(ss,CFG.ABAS.HORAS_EXTRAS,['ID','Data','ID Funcionário','Funcionário','Minutos','Motivo','Salário base','Horas semanais','Divisor mensal','Valor hora','Valor total','Status','Data de cadastro']);criarAba_(ss,CFG.ABAS.LOG,['Data','Tipo','ID','Motivo','Usuário']);criarAba_(ss,CFG.ABAS.COMISSOES,['ID','Data do pagamento','Vendedora','Valor pago','Referência','Registrado por']);migrarProdutosExistentes_(ss.getSheetByName(CFG.ABAS.PRODUTOS));const sh=ss.getSheetByName(CFG.ABAS.USUARIOS);if(sh.getLastRow()<2)sh.appendRow([uid_('USR'),'Administrador','admin',hash_('123456'),'Administrador','Ativo','','',new Date()]);return 'Sistema configurado. Login inicial: admin / 123456';}
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
function doPost(e){try{const d=JSON.parse(e.postData.contents||'{}'),a=d.acao;if(a==='login')return json_(login_(d));const u=auth_(d.token);const mapa={verificarToken:()=>({ok:true,usuario:publicUser_(u)}),dashboard:()=>dashboard_(u),listarProdutos:()=>listarProdutos_(u),salvarProduto:()=>salvarProduto_(u,d),alterarStatusProduto:()=>alterarStatusProduto_(u,d),alterarPromocaoProduto:()=>alterarPromocaoProduto_(u,d),excluirProduto:()=>excluirProduto_(u,d),cadastrarVenda:()=>cadastrarVenda_(u,d),listarVendas:()=>listarVendas_(u),listarComissoes:()=>listarComissoes_(u),alterarStatusComissao:()=>alterarStatusComissao_(u,d),registrarPagamentoComissao:()=>registrarPagamentoComissao_(u,d),listarContasReceber:()=>listarContasReceber_(u),alterarStatusParcela:()=>alterarStatusParcela_(u,d),baixarParcela:()=>baixarParcela_(u,d),registrarRecebimento:()=>registrarRecebimento_(u,d),excluirVenda:()=>excluirVenda_(u,d),listarUsuarios:()=>listarUsuarios_(u),salvarUsuario:()=>salvarUsuario_(u,d),excluirUsuario:()=>excluirUsuario_(u,d),listarHorasExtras:()=>listarHorasExtras_(u),registrarHoraExtra:()=>registrarHoraExtra_(u,d),alterarStatusHoraExtra:()=>alterarStatusHoraExtra_(u,d),excluirHoraExtra:()=>excluirHoraExtra_(u,d),listarCatalogoRevendedor:()=>listarCatalogoRevendedor_(u)};if(!mapa[a])throw new Error('Ação inválida.');return json_(mapa[a]())}catch(err){return json_({ok:false,erro:msg_(err)})}}
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
function auth_(token){if(!token)throw new Error('Sessão expirada.');const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS),u=objetos_(sh).find(x=>String(x.Token||'')===String(token)&&x.Status==='Ativo');if(!u||new Date(u['Validade token']).getTime()<Date.now())throw new Error('Sessão expirada.');return u}function admin_(u){if(u.Perfil!=='Administrador')throw new Error('Acesso permitido somente ao administrador.')}function gerenciaProdutos_(u){if(!['Administrador','Vendedora','Vendedor'].includes(String(u.Perfil||'')))throw new Error('Você não tem permissão para gerenciar produtos.')}function publicUser_(u){return{id:u.ID,nome:u.Nome,usuario:u['Usuário'],perfil:u.Perfil}}
function listarProdutos_(){const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);return{ok:true,produtos:objetos_(sh).map(produtoNormalizado_).filter(x=>x.nome)}}
function salvarProduto_(u,d){
  gerenciaProdutos_(u);
  const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS);
  const todos=objetos_(sh);
  let r=Number(d.linha)>0?todos.find(x=>Number(x._linha)===Number(d.linha)):null;
  if(!r&&d.id)r=todos.find(x=>String(valorCampo_(x,['ID'])||'')===String(d.id));
  let fotos=r?fotos_(r):[];
  if(Array.isArray(d.fotosUpload)&&d.fotosUpload.length)fotos=salvarFotos_(d.fotosUpload,d.nome);
  const id=r?(valorCampo_(r,['ID'])||uid_('PROD')):uid_('PROD');
  const obj={'Produto':d.nome,'Categoria':d.categoria,'Preço':num_(d.preco),'Preço de custo':num_(d.precoCusto),'Preço revenda':calcularPrecoRevenda_(num_(d.preco),d.categoria),'Classificação':d.classificacao||'A','Descrição':d.descricao||'','Comentário revendedor':d.comentarioRevendedor||'','Link da Imagem':fotos[0]||'','Promoção':d.promocao!==undefined&&d.promocao!==''?d.promocao:(r?(valorCampo_(r,['Promoção','Promocao'])||'Não'):'Não'),'WhatsApp':r?(valorCampo_(r,['WhatsApp','Whatsapp'])||''):'','ID':id,'Marca':d.marca||'','Modelo':d.modelo||'','Quantidade':num_(d.quantidade),'Status':d.status||'Publicado','Fotos':JSON.stringify(fotos),'Data de cadastro':r?(valorCampo_(r,['Data de cadastro'])||new Date()):new Date()};
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
function alterarStatusProduto_(u,d){gerenciaProdutos_(u);const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS),todos=objetos_(sh);let r=todos.find(x=>String(x.ID||'')===String(d.id||''));if(!r&&Number(d.linha)>0)r=todos.find(x=>Number(x._linha)===Number(d.linha));if(!r)throw new Error('Produto não encontrado. Atualize a página e tente novamente.');const status=statusNormalizado_(d.status);atualizarLinha_(sh,r._linha,{Status:status});SpreadsheetApp.flush();return{ok:true,status,linha:r._linha}}
function alterarPromocaoProduto_(u,d){
  gerenciaProdutos_(u);
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
function excluirProduto_(u,d){gerenciaProdutos_(u);const sh=planilha_().getSheetByName(CFG.ABAS.PRODUTOS),r=objetos_(sh).find(x=>x.ID===d.id);if(!r)throw new Error('Produto não encontrado.');sh.deleteRow(r._linha);log_(u,'Produto',d.id,d.motivo);return{ok:true}}

function apenasDigitos_(v){return String(v||'').replace(/\D/g,'')}
function cpfValido_(cpf){cpf=apenasDigitos_(cpf);if(cpf.length!==11||/^(\d)\1{10}$/.test(cpf))return false;let soma=0;for(let i=0;i<9;i++)soma+=Number(cpf[i])*(10-i);let d=(soma*10)%11;if(d===10)d=0;if(d!==Number(cpf[9]))return false;soma=0;for(let i=0;i<10;i++)soma+=Number(cpf[i])*(11-i);d=(soma*10)%11;if(d===10)d=0;return d===Number(cpf[10])}
function cnpjValido_(cnpj){cnpj=apenasDigitos_(cnpj);if(cnpj.length!==14||/^(\d)\1{13}$/.test(cnpj))return false;function calc(p){let soma=0;for(let i=0;i<p.length;i++)soma+=Number(cnpj[i])*p[i];const r=soma%11;return r<2?0:11-r}const d1=calc([5,4,3,2,9,8,7,6,5,4,3,2]);if(d1!==Number(cnpj[12]))return false;return calc([6,5,4,3,2,9,8,7,6,5,4,3,2])===Number(cnpj[13])}
function documentoValido_(v){const d=apenasDigitos_(v);return !d||(d.length===11?cpfValido_(d):d.length===14&&cnpjValido_(d))}
function telefoneValido_(v){const d=apenasDigitos_(v);return !d||((d.length===10||d.length===11)&&d[0]!=='0'&&d[1]!=='0')}

function cadastrarVenda_(u,d){
  const ss=planilha_();
  const ps=ss.getSheetByName(CFG.ABAS.PRODUTOS);
  const vs=criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria','Entrada','Saldo a receber','Parcelas JSON','Status recebimento','Recebimentos JSON','Status comissão','Comissão revisada por','Data revisão comissão']);
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
  const obj={'ID Venda':id,'Data da venda':data,'Data de cadastro':new Date(),'Cliente':d.cliente,'CPF/CNPJ':documento,'Telefone':telefone,'ID Produto':idProduto,'Produto':produto,'Quantidade':qtd,'Valor da venda':valor,'Custo unitário':custoUnitario,'Custo total':custoTotal,'Lucro':lucro,'Forma de pagamento':d.pagamento,'Garantia dias':num_(d.garantia),'Observações':d.observacoes||'','Vendedora':u.Nome,'Origem do produto':avulsa?'Venda avulsa':'Estoque','Marca':marca,'Modelo':modelo,'Categoria':categoria,'Entrada':num_(d.entrada),'Saldo a receber':Math.max(0,valor-num_(d.entrada)),'Parcelas JSON':'[]','Status recebimento':Math.max(0,valor-num_(d.entrada))>0?'Pendente':'Pago','Recebimentos JSON':'[]','Status comissão':'Pendente','Comissão revisada por':'','Data revisão comissão':''};
  appendObj_(vs,obj);
  if(p){const pn=produtoNormalizado_(p),nova=pn.quantidade-qtd;atualizarLinha_(ps,p._linha,{Quantidade:nova,Status:nova<=0?'Oculto':pn.status});}
  return{ok:true,venda:{id,cliente:d.cliente,documento,telefone,produto,quantidade:qtd,valor,custoUnitario,custoTotal,lucro,pagamento:d.pagamento,garantia:num_(d.garantia),observacoes:d.observacoes||'',data,vendedora:u.Nome,origem:avulsa?'Venda avulsa':'Estoque',entrada:num_(d.entrada),saldo:Math.max(0,valor-num_(d.entrada)),parcelas:[],recebimentos:[]}};
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
    let parcelas=[];try{parcelas=JSON.parse(x['Parcelas JSON']||'[]')}catch(e){}let recebimentos=[];try{recebimentos=JSON.parse(x['Recebimentos JSON']||'[]')}catch(e){}const totalParcelas=parcelas.length,parcelasPagas=parcelas.filter(p=>String(p.status||'Pendente')==='Pago').length;return{id:x['ID Venda'],data:fmtData_(x['Data da venda']),cliente:x.Cliente,documento:x['CPF/CNPJ'],telefone:x.Telefone,idProduto,produto:x.Produto,quantidade,valor,custoUnitario,custoTotal,lucro,pagamento:x['Forma de pagamento'],garantia:num_(x['Garantia dias']),observacoes:x['Observações'],vendedora:x.Vendedora,origem:x['Origem do produto']||((x['ID Produto'])?'Estoque':'Venda avulsa'),entrada:num_(x.Entrada),saldo:num_(x['Saldo a receber']),parcelas,totalParcelas,parcelasPagas,statusRecebimento:x['Status recebimento']||'',recebimentos,statusComissao:(()=>{const s=String(x['Status comissão']||'').trim();return ['Aprovada','Recusada'].includes(s)?s:'Pendente'})(),comissaoRevisadaPor:x['Comissão revisada por']||'',dataRevisaoComissao:fmtData_(x['Data revisão comissão'])};
  });
  return{ok:true,vendas};
}

function listarComissoes_(u){
  const percentual=0.02;
  const hoje=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const mes=hoje.slice(0,7);
  let vendas=listarVendas_().vendas;
  if(u.Perfil!=='Administrador')vendas=vendas.filter(v=>String(v.vendedora||'').trim().toLowerCase()===String(u.Nome||'').trim().toLowerCase());
  const itens=vendas.map(v=>{
    const valor=num_(v.valor),recebido=Math.max(0,valor-num_(v.saldo)),pendente=Math.max(0,num_(v.saldo));
    const status=['Aprovada','Recusada'].includes(String(v.statusComissao||''))?String(v.statusComissao):'Pendente';
    const comissaoGerada=Number((valor*percentual).toFixed(2));
    return {...v,recebido,pendente,statusComissao:status,comissaoGerada,comissaoTotal:comissaoGerada,comissaoPendente:status==='Pendente'?comissaoGerada:0,comissaoRecusada:status==='Recusada'?comissaoGerada:0};
  });
  const sh=criarAba_(planilha_(),CFG.ABAS.COMISSOES,['ID','Data do pagamento','Vendedora','Valor pago','Referência','Registrado por']);
  let pagamentos=objetos_(sh).map(x=>({id:x.ID,data:fmtData_(x['Data do pagamento']),dataIso:isoData_(x['Data do pagamento']),vendedora:x.Vendedora,valor:num_(x['Valor pago']),referencia:x['Referência']||'',registradoPor:x['Registrado por']||''}));
  if(u.Perfil!=='Administrador')pagamentos=pagamentos.filter(p=>String(p.vendedora||'').trim().toLowerCase()===String(u.Nome||'').trim().toLowerCase());
  const pagoPorVendedor={};
  pagamentos.forEach(p=>{const nome=String(p.vendedora||'Sem vendedor').trim().toLowerCase();pagoPorVendedor[nome]=(pagoPorVendedor[nome]||0)+num_(p.valor)});
  const restantePago={...pagoPorVendedor};
  itens.sort((a,b)=>String(a.data||'').localeCompare(String(b.data||''))).forEach(v=>{
    const chave=String(v.vendedora||'Sem vendedor').trim().toLowerCase();
    const disponivel=v.statusComissao==='Aprovada'?num_(v.comissaoGerada):0;
    const abatido=Math.min(disponivel,Math.max(0,restantePago[chave]||0));
    if(disponivel>0)restantePago[chave]=Math.max(0,(restantePago[chave]||0)-abatido);
    v.comissaoPaga=Number(abatido.toFixed(2));
    v.comissaoLiberada=Number(Math.max(0,disponivel-abatido).toFixed(2));
  });
  const totalGerado=itens.reduce((s,v)=>s+num_(v.comissaoGerada),0);
  const totalPendente=itens.filter(v=>v.statusComissao==='Pendente').reduce((s,v)=>s+num_(v.comissaoGerada),0);
  const totalPago=pagamentos.reduce((s,p)=>s+num_(p.valor),0);
  const saldoPagar=itens.reduce((s,v)=>s+num_(v.comissaoLiberada),0);
  const liberadoHoje=itens.filter(v=>v.statusComissao==='Aprovada'&&isoData_(v.data)===hoje).reduce((s,v)=>s+num_(v.comissaoLiberada),0);
  const liberadoMes=itens.filter(v=>v.statusComissao==='Aprovada'&&isoData_(v.data).slice(0,7)===mes).reduce((s,v)=>s+num_(v.comissaoLiberada),0);
  const vendedores={};
  itens.forEach(v=>{
    const nome=v.vendedora||'Sem vendedor';
    if(!vendedores[nome])vendedores[nome]={vendedora:nome,gerada:0,liberada:0,pendente:0,paga:0};
    vendedores[nome].gerada+=v.statusComissao==='Aprovada'?num_(v.comissaoGerada):0;
    vendedores[nome].liberada+=num_(v.comissaoLiberada);
    vendedores[nome].pendente+=v.statusComissao==='Pendente'?num_(v.comissaoGerada):0;
  });
  pagamentos.forEach(p=>{const nome=p.vendedora||'Sem vendedor';if(!vendedores[nome])vendedores[nome]={vendedora:nome,gerada:0,liberada:0,pendente:0,paga:0};vendedores[nome].paga+=num_(p.valor)});
  const resumoVendedores=Object.values(vendedores).map(x=>({...x,saldoPagar:Number(Math.max(0,x.liberada).toFixed(2))})).filter(x=>x.saldoPagar>0.009);
  pagamentos.sort((a,b)=>String(b.dataIso||'').localeCompare(String(a.dataIso||'')));
  itens.sort((a,b)=>isoData_(b.data).localeCompare(isoData_(a.data)));
  return{ok:true,percentual:2,totalHoje:liberadoHoje,totalMes:liberadoMes,totalLiberado:Number(saldoPagar.toFixed(2)),totalGerado:Number(totalGerado.toFixed(2)),totalPendente:Number(totalPendente.toFixed(2)),totalPago:Number(totalPago.toFixed(2)),saldoPagar:Number(saldoPagar.toFixed(2)),vendasMes:itens.filter(v=>isoData_(v.data).slice(0,7)===mes).length,itens,pagamentos,resumoVendedores};
}
function alterarStatusComissao_(u,d){
  admin_(u);
  const status=String(d.status||'').trim();
  if(!['Aprovada','Recusada','Pendente'].includes(status))throw new Error('Status de comissão inválido.');
  const sh=criarAba_(planilha_(),CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria','Entrada','Saldo a receber','Parcelas JSON','Status recebimento','Recebimentos JSON','Status comissão','Comissão revisada por','Data revisão comissão']);
  const venda=objetos_(sh).find(x=>String(x['ID Venda']||'')===String(d.vendaId||''));
  if(!venda)throw new Error('Venda não encontrada.');
  atualizarLinha_(sh,venda._linha,{'Status comissão':status,'Comissão revisada por':u.Nome,'Data revisão comissão':new Date()});
  return{ok:true,status};
}
function registrarPagamentoComissao_(u,d){
  admin_(u);
  const valor=num_(d.valor);if(valor<=0)throw new Error('Informe um valor válido.');
  const vendedora=String(d.vendedora||'').trim();if(!vendedora)throw new Error('Informe a vendedora.');
  const dados=listarComissoes_({Perfil:'Administrador',Nome:''});
  const resumo=(dados.resumoVendedores||[]).find(x=>String(x.vendedora).trim().toLowerCase()===vendedora.toLowerCase());
  const disponivel=resumo?num_(resumo.saldoPagar):0;if(valor>disponivel+0.01)throw new Error('O valor é maior que a comissão disponível.');
  const sh=criarAba_(planilha_(),CFG.ABAS.COMISSOES,['ID','Data do pagamento','Vendedora','Valor pago','Referência','Registrado por']);
  appendObj_(sh,{'ID':uid_('COM'),'Data do pagamento':new Date(),'Vendedora':vendedora,'Valor pago':valor,'Referência':d.referencia||'Pagamento de comissão','Registrado por':u.Nome});
  return{ok:true};
}

function historicoRecebimentosVenda_(v){
  const existentes=Array.isArray(v.recebimentos)?v.recebimentos.filter(r=>num_(r.valor)>0):[];
  if(existentes.length)return existentes;
  const parcelas=Array.isArray(v.parcelas)?v.parcelas:[];
  const pagas=parcelas.filter(p=>String(p.status||'').toLowerCase()==='pago'&&num_(p.valor)>0).map((p,i)=>({id:'LEG-PARC-'+(p.numero||i+1),data:p.dataPagamento||p.vencimento||v.data,valor:Number(num_(p.valor).toFixed(2)),observacao:'Pagamento registrado no sistema anterior',registradoPor:'Migração automática'}));
  if(pagas.length)return pagas;
  const pagoAposEntrada=Math.max(0,num_(v.valor)-num_(v.entrada)-num_(v.saldo));
  if(pagoAposEntrada>0.009)return[{id:'LEG-SALDO-'+v.id,data:v.data,valor:Number(pagoAposEntrada.toFixed(2)),observacao:'Pagamento anterior recuperado pelo saldo da venda',registradoPor:'Migração automática'}];
  return[];
}
function listarContasReceber_(u){
  const vendas=listarVendas_().vendas;
  const hoje=new Date();hoje.setHours(0,0,0,0);
  const contas=vendas.filter(v=>num_(v.saldo)>0.009).map(v=>{
    const recebimentos=historicoRecebimentosVenda_(v);
    const datas=recebimentos.map(r=>isoData_(r.data)).filter(Boolean).sort();
    const ultimoIso=datas.length?datas[datas.length-1]:(num_(v.entrada)>0?isoData_(v.data):'');
    let diasSemPagamento=0;
    if(ultimoIso){const partes=ultimoIso.split('-').map(Number),d=new Date(partes[0],partes[1]-1,partes[2]);d.setHours(0,0,0,0);diasSemPagamento=Math.max(0,Math.floor((hoje-d)/86400000));}
    const totalPago=Math.max(0,num_(v.valor)-num_(v.saldo));
    const quantidadePagamentos=recebimentos.length+(num_(v.entrada)>0?1:0);
    return{vendaId:v.id,cliente:v.cliente,documento:v.documento,telefone:v.telefone,produto:v.produto,valorVenda:num_(v.valor),entrada:num_(v.entrada),saldo:num_(v.saldo),totalPago:Number(totalPago.toFixed(2)),quantidadePagamentos,dataVenda:v.data,ultimoPagamento:ultimoIso?fmtData_(ultimoIso):'',diasSemPagamento,vendedora:v.vendedora,status:v.statusRecebimento||'Pendente',recebimentos};
  });
  contas.sort((a,b)=>Number(b.diasSemPagamento||0)-Number(a.diasSemPagamento||0)||String(b.dataVenda||'').localeCompare(String(a.dataVenda||'')));
  return{ok:true,contas,totalReceber:contas.reduce((s,c)=>s+num_(c.saldo),0),clientesPendentes:new Set(contas.map(c=>String(c.vendaId))).size};
}
function registrarRecebimento_(u,d){
  const sh=planilha_().getSheetByName(CFG.ABAS.VENDAS),r=objetos_(sh).find(x=>String(x['ID Venda'])===String(d.vendaId));
  if(!r)throw new Error('Venda não encontrada.');
  const valor=num_(d.valor);if(valor<=0)throw new Error('Informe um valor válido.');
  const saldoAtual=Math.max(0,num_(r['Saldo a receber']));if(valor>saldoAtual+0.01)throw new Error('O pagamento não pode ser maior que o saldo devedor.');
  let recebimentos=[];try{recebimentos=JSON.parse(r['Recebimentos JSON']||'[]')}catch(e){}
  if(!Array.isArray(recebimentos))recebimentos=[];
  if(!recebimentos.length){
    let parcelasAntigas=[];try{parcelasAntigas=JSON.parse(r['Parcelas JSON']||'[]')}catch(e){}
    recebimentos=historicoRecebimentosVenda_({id:r['ID Venda'],data:fmtData_(r['Data da venda']),valor:num_(r['Valor da venda']),entrada:num_(r.Entrada),saldo:saldoAtual,parcelas:parcelasAntigas,recebimentos:[]});
  }
  recebimentos.push({id:uid_('REC'),data:d.dataPagamento||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'),valor:Number(valor.toFixed(2)),observacao:String(d.observacao||''),registradoPor:u.Nome});
  const saldo=Number(Math.max(0,saldoAtual-valor).toFixed(2));
  atualizarLinha_(sh,r._linha,{'Recebimentos JSON':JSON.stringify(recebimentos),'Saldo a receber':saldo,'Status recebimento':saldo>0?'Pendente':'Pago'});
  return{ok:true,saldo,recebimentos};
}
function alterarStatusParcela_(u,d){
  const sh=planilha_().getSheetByName(CFG.ABAS.VENDAS),r=objetos_(sh).find(x=>String(x['ID Venda'])===String(d.vendaId));if(!r)throw new Error('Venda não encontrada.');
  let parcelas=[];try{parcelas=JSON.parse(r['Parcelas JSON']||'[]')}catch(e){}
  if(!parcelas.length&&num_(r['Saldo a receber'])>0)parcelas=[{numero:1,vencimento:isoData_(r['Data da venda']),valor:num_(r['Saldo a receber']),status:'Pendente',dataPagamento:''}];
  const indice=Number(d.indice);if(indice<0||indice>=parcelas.length)throw new Error('Parcela não encontrada.');
  const paga=(d.paga===true||String(d.paga).toLowerCase()==='true'||String(d.paga)==='1');parcelas[indice].status=paga?'Pago':'Pendente';parcelas[indice].dataPagamento=paga?Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'):'';
  const saldo=parcelas.filter(p=>String(p.status||'Pendente')!=='Pago').reduce((s,p)=>s+num_(p.valor),0);
  atualizarLinha_(sh,r._linha,{'Parcelas JSON':JSON.stringify(parcelas),'Saldo a receber':saldo,'Status recebimento':saldo>0?'Pendente':'Pago'});return{ok:true,saldo};
}

function baixarParcela_(u,d){const sh=planilha_().getSheetByName(CFG.ABAS.VENDAS),r=objetos_(sh).find(x=>String(x['ID Venda'])===String(d.vendaId));if(!r)throw new Error('Venda não encontrada.');let parcelas=[];try{parcelas=JSON.parse(r['Parcelas JSON']||'[]')}catch(e){}const p=parcelas.find((x,i)=>Number(x.numero||i+1)===Number(d.numero));if(!p)throw new Error('Parcela não encontrada.');p.status='Pago';p.dataPagamento=d.dataPagamento||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');const saldo=parcelas.filter(x=>x.status!=='Pago').reduce((a,x)=>a+num_(x.valor),0);atualizarLinha_(sh,r._linha,{'Parcelas JSON':JSON.stringify(parcelas),'Saldo a receber':saldo,'Status recebimento':saldo>0?'Pendente':'Pago'});return{ok:true,saldo};}
function excluirVenda_(u,d){admin_(u);const ss=planilha_(),vs=ss.getSheetByName(CFG.ABAS.VENDAS),r=objetos_(vs).find(x=>x['ID Venda']===d.id);if(!r)throw new Error('Venda não encontrada.');const ps=ss.getSheetByName(CFG.ABAS.PRODUTOS),p=objetos_(ps).find(x=>String(x.ID)===String(r['ID Produto']));if(p)atualizarLinha_(ps,p._linha,{Quantidade:produtoNormalizado_(p).quantidade+num_(r.Quantidade)});vs.deleteRow(r._linha);log_(u,'Venda',d.id,d.motivo);return{ok:true}}
function dashboard_(){const hoje=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'),mes=hoje.slice(0,7),v=listarVendas_().vendas;const vh=v.filter(x=>isoData_(x.data)===hoje),vm=v.filter(x=>isoData_(x.data).slice(0,7)===mes);const custoHoje=vh.reduce((s,x)=>s+num_(x.custoTotal),0),custoMes=vm.reduce((s,x)=>s+num_(x.custoTotal),0),faturamentoHoje=vh.reduce((s,x)=>s+x.valor,0),faturamentoMes=vm.reduce((s,x)=>s+x.valor,0),totalReceber=v.reduce((s,x)=>s+num_(x.saldo),0);return{ok:true,faturamentoHoje,faturamentoMes,custoHoje,custoMes,lucroHoje:faturamentoHoje-custoHoje,lucroMes:faturamentoMes-custoMes,totalReceber,vendasMes:vm.length,estoque:listarProdutos_().produtos.reduce((s,x)=>s+x.quantidade,0),ultimasVendas:v.slice(0,8)}}
function listarUsuarios_(u){
  admin_(u);
  const sh=criarAba_(planilha_(),CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro','Salário mensal','Horas semanais']);
  return{ok:true,usuarios:objetos_(sh).map(x=>({id:x.ID,nome:x.Nome,usuario:x['Usuário'],perfil:x.Perfil,status:x.Status,salario:num_(x['Salário mensal']),horasSemanais:num_(x['Horas semanais'])}))};
}
function salvarUsuario_(u,d){
  admin_(u);
  const sh=criarAba_(planilha_(),CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro','Salário mensal','Horas semanais']);
  const todos=objetos_(sh),r=d.id?todos.find(x=>x.ID===d.id):null;
  if(todos.some(x=>String(x['Usuário']).toLowerCase()===String(d.usuario).toLowerCase()&&x.ID!==d.id))throw new Error('Esse usuário já existe.');
  const perfil=d.perfil||'Vendedora',salario=num_(d.salario),horasSemanais=num_(d.horasSemanais);
  if(perfil==='Vendedora'&&(salario<=0||horasSemanais<=0))throw new Error('Informe o salário e a carga horária semanal da vendedora.');
  const obj={'ID':r?r.ID:uid_('USR'),'Nome':d.nome,'Usuário':d.usuario,'Senha Hash':d.senha?hash_(d.senha):(r?r['Senha Hash']:''),'Perfil':perfil,'Status':d.status||'Ativo','Token':r?r.Token:'','Validade token':r?r['Validade token']:'','Data de cadastro':r?r['Data de cadastro']:new Date(),'Salário mensal':salario,'Horas semanais':horasSemanais};
  if(r)atualizarLinha_(sh,r._linha,obj);else appendObj_(sh,obj);
  return{ok:true};
}
function abaHorasExtras_(){
  return criarAba_(planilha_(),CFG.ABAS.HORAS_EXTRAS,['ID','Data','ID Funcionário','Funcionário','Minutos','Motivo','Salário base','Horas semanais','Divisor mensal','Valor hora','Valor total','Status','Data de cadastro']);
}
function dadosFolhaUsuario_(id){
  const sh=criarAba_(planilha_(),CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro','Salário mensal','Horas semanais']);
  const usuario=objetos_(sh).find(x=>String(x.ID||'')===String(id||''));
  if(!usuario)throw new Error('Funcionário não encontrado.');
  const salario=num_(usuario['Salário mensal']),horasSemanais=num_(usuario['Horas semanais']);
  if(salario<=0||horasSemanais<=0)throw new Error('Seu salário e sua carga horária semanal ainda não foram cadastrados pelo administrador.');
  const divisorMensal=Number((horasSemanais*5).toFixed(2));
  const valorHora=divisorMensal>0?Number((salario/divisorMensal).toFixed(6)):0;
  return{usuario,salario,horasSemanais,divisorMensal,valorHora};
}
function registrarHoraExtra_(u,d){
  if(u.Perfil==='Revendedor')throw new Error('Recurso não disponível para revendedor.');
  const folha=dadosFolhaUsuario_(u.ID);
  const horas=Math.max(0,Math.floor(num_(d.horas))),minutosParte=Math.max(0,Math.floor(num_(d.minutos)));
  if(minutosParte>59)throw new Error('Os minutos devem estar entre 0 e 59.');
  const minutos=horas*60+minutosParte;
  if(minutos<=0)throw new Error('Informe a quantidade de horas ou minutos extras.');
  if(minutos>24*60)throw new Error('O lançamento não pode ultrapassar 24 horas.');
  const hoje=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const data=isoData_(d.data||hoje);
  if(!/^\d{4}-\d{2}-\d{2}$/.test(data))throw new Error('Informe uma data válida.');
  if(data>hoje)throw new Error('Não é possível cadastrar hora extra em data futura.');
  const motivo=String(d.motivo||'').trim();if(!motivo)throw new Error('Informe o motivo da hora extra.');
  const valorTotal=Number((folha.valorHora*(minutos/60)).toFixed(2)),id=uid_('HEX');
  appendObj_(abaHorasExtras_(),{'ID':id,'Data':data,'ID Funcionário':u.ID,'Funcionário':u.Nome,'Minutos':minutos,'Motivo':motivo,'Salário base':folha.salario,'Horas semanais':folha.horasSemanais,'Divisor mensal':folha.divisorMensal,'Valor hora':folha.valorHora,'Valor total':valorTotal,'Status':'Pendente','Data de cadastro':new Date()});
  return{ok:true,registro:{id,data:fmtData_(data),idFuncionario:u.ID,funcionario:u.Nome,minutos,motivo,valorHora:folha.valorHora,valorTotal,status:'Pendente'}};
}
function listarHorasExtras_(u){
  if(u.Perfil==='Revendedor')return{ok:true,registros:[],minutosPendentes:0,valorPendente:0,minutosPagos:0,valorPago:0,valorTotal:0};
  let registros=objetos_(abaHorasExtras_());
  if(u.Perfil!=='Administrador')registros=registros.filter(x=>String(x['ID Funcionário']||'')===String(u.ID||''));
  const saida=registros.reverse().map(x=>({id:x.ID,data:fmtData_(x.Data),dataIso:isoData_(x.Data),idFuncionario:x['ID Funcionário'],funcionario:x.Funcionário,minutos:num_(x.Minutos),motivo:x.Motivo||'',salarioBase:num_(x['Salário base']),horasSemanais:num_(x['Horas semanais']),divisorMensal:num_(x['Divisor mensal']),valorHora:num_(x['Valor hora']),valorTotal:num_(x['Valor total']),status:x.Status||'Pendente'}));
  const pendentes=saida.filter(x=>x.status!=='Pago'),pagos=saida.filter(x=>x.status==='Pago');
  return{ok:true,registros:saida,minutosPendentes:pendentes.reduce((a,x)=>a+x.minutos,0),valorPendente:Number(pendentes.reduce((a,x)=>a+x.valorTotal,0).toFixed(2)),minutosPagos:pagos.reduce((a,x)=>a+x.minutos,0),valorPago:Number(pagos.reduce((a,x)=>a+x.valorTotal,0).toFixed(2)),valorTotal:Number(saida.reduce((a,x)=>a+x.valorTotal,0).toFixed(2))};
}
function alterarStatusHoraExtra_(u,d){
  admin_(u);const sh=abaHorasExtras_(),r=objetos_(sh).find(x=>String(x.ID||'')===String(d.id||''));if(!r)throw new Error('Lançamento não encontrado.');const status=String(d.status||'')==='Pago'?'Pago':'Pendente';atualizarLinha_(sh,r._linha,{'Status':status});return{ok:true,status};
}
function excluirHoraExtra_(u,d){
  const sh=abaHorasExtras_(),r=objetos_(sh).find(x=>String(x.ID||'')===String(d.id||''));if(!r)throw new Error('Lançamento não encontrado.');const dono=String(r['ID Funcionário']||'')===String(u.ID||'');if(u.Perfil!=='Administrador'&&!dono)throw new Error('Você não pode excluir este lançamento.');if(u.Perfil!=='Administrador'&&String(r.Status||'Pendente')==='Pago')throw new Error('Uma hora extra já paga só pode ser alterada pelo administrador.');sh.deleteRow(r._linha);return{ok:true};
}
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
    comentarioRevendedor:String(valorCampo_(x,['Comentário revendedor','Comentario revendedor'])||''),
    precoRevenda:calcularPrecoRevenda_(num_(valorCampo_(x,['Preço','Preco'])),String(valorCampo_(x,['Categoria'])||'')),
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
  'Produto','Categoria','Preço','Preço revenda','Preço de custo','Classificação','Descrição','Comentário revendedor','Link da Imagem','Promoção',
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
  criarAba_(ss,CFG.ABAS.VENDAS,['ID Venda','Data da venda','Data de cadastro','Cliente','CPF/CNPJ','Telefone','ID Produto','Produto','Quantidade','Valor da venda','Custo unitário','Custo total','Lucro','Forma de pagamento','Garantia dias','Observações','Vendedora','Origem do produto','Marca','Modelo','Categoria','Entrada','Saldo a receber','Parcelas JSON','Status recebimento','Recebimentos JSON','Status comissão','Comissão revisada por','Data revisão comissão']);
  criarAba_(ss,CFG.ABAS.USUARIOS,['ID','Nome','Usuário','Senha Hash','Perfil','Status','Token','Validade token','Data de cadastro','Salário mensal','Horas semanais']);
  criarAba_(ss,CFG.ABAS.HORAS_EXTRAS,['ID','Data','ID Funcionário','Funcionário','Minutos','Motivo','Salário base','Horas semanais','Divisor mensal','Valor hora','Valor total','Status','Data de cadastro']);
  criarAba_(ss,CFG.ABAS.LOG,['Data','Tipo','ID','Motivo','Usuário']);
  criarAba_(ss,CFG.ABAS.COMISSOES,['ID','Data do pagamento','Vendedora','Valor pago','Referência','Registrado por']);
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
    const precoBase=num_(ultimo(linha,['Preço','Preco']));
    const categoriaBase=primeiro(linha,['Categoria'])||'outros';
    saida.push([
      produto,
      categoriaBase,
      precoBase,
      calcularPrecoRevenda_(precoBase,categoriaBase),
      num_(ultimo(linha,['Preço de custo','Preco de custo','Custo'])),
      primeiro(linha,['Classificação','Classificacao'])||'A',
      primeiro(linha,['Descrição','Descricao']),
      primeiro(linha,['Comentário revendedor','Comentario revendedor']),
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
  gerenciaProdutos_(u);
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


function categoriaLinhaPesada_(categoria){
  const c=normalizarChave_(categoria).replace(/-/g,' ');
  const pesadas=['geladeira','lavadora','fogao','micro ondas','microondas','ar condicionado','frigobar','maquina de lavar','lava e seca','freezer'];
  return pesadas.some(x=>c===x||c.includes(x));
}
function calcularPrecoRevenda_(preco,categoria){
  const p=num_(preco);if(p<=0)return 0;
  const desconto=categoriaLinhaPesada_(categoria)?0.11:0.18;
  return Math.round((p*(1-desconto))*100)/100;
}
function listarCatalogoRevendedor_(u){
  if(String(u.Perfil)!=='Revendedor'&&String(u.Perfil)!=='Administrador')throw new Error('Acesso restrito aos revendedores.');
  const produtos=listarProdutos_().produtos.filter(x=>x.status==='Publicado'&&Number(x.quantidade)>0&&Number(x.preco)>0).map(x=>({
    id:x.id,nome:x.nome,categoria:x.categoria,marca:x.marca,modelo:x.modelo,descricao:x.descricao,classificacao:x.classificacao,
    comentarioRevendedor:x.comentarioRevendedor,fotos:x.fotos,quantidade:x.quantidade,precoNormal:x.preco,
    precoRevenda:calcularPrecoRevenda_(x.preco,x.categoria),
    descontoPercentual:categoriaLinhaPesada_(x.categoria)?11:18
  }));
  return {ok:true,produtos};
}
function excluirUsuario_(u,d){
  admin_(u);const sh=planilha_().getSheetByName(CFG.ABAS.USUARIOS);const todos=objetos_(sh);const r=todos.find(x=>String(x.ID)===String(d.id));
  if(!r)throw new Error('Cadastro não encontrado.');if(String(r.ID)===String(u.ID))throw new Error('Você não pode excluir o próprio usuário.');
  sh.deleteRow(r._linha);return {ok:true};
}
