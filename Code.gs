function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Produtos');

    if (!sheet) {
      return responder(e, {
        ok: false,
        erro: 'A aba Produtos não foi encontrada.',
        html: '<div class="sem-produtos">A aba Produtos não foi encontrada.</div>'
      });
    }

    const values = sheet.getDataRange().getDisplayValues();

    if (values.length < 2) {
      return responder(e, {
        ok: true,
        html: '<div class="sem-produtos">Nenhum produto cadastrado.</div>'
      });
    }

    const rows = values.slice(1).filter(row => String(row[0] || '').trim());

    const html = rows.map((row, index) => {
      const produtoOriginal = String(row[0] || '').trim();
      const categoriaOriginal = String(row[1] || '').trim();
      const precoTexto = String(row[2] || '').trim();
      const classificacaoOriginal = String(row[3] || '').trim();
      const descricaoOriginal = String(row[4] || '').trim();
      const imagemOriginal = String(row[5] || '').trim();
      const promocaoOriginal = String(row[6] || '').trim();

      const produto = escapeHtml(produtoOriginal);
      const categoria = escapeHtml(categoriaOriginal);
      const classificacao = escapeHtml(classificacaoOriginal);
      const descricao = escapeHtml(descricaoOriginal);
      const imagem = normalizarImagem(imagemOriginal);
      const promocao = ehPromocao(promocaoOriginal) ? 'sim' : 'nao';
      const preco = normalizarPreco(precoTexto);

      const mensagem = encodeURIComponent(
        'Olá! Tenho interesse no produto: ' + produtoOriginal +
        (classificacaoOriginal ? ' - Classificação ' + classificacaoOriginal : '')
      );

      return `
        <article class="produto ${promocao === 'sim' ? 'produto-promocao' : ''}"
          data-categoria="${categoria}"
          data-categoria-nome="${categoria}"
          data-promocao="${promocao}"
          data-oferta="${promocao}">
          <div class="produto-imagem">
            <img src="${imagem}" alt="${produto}" loading="${index < 4 ? 'eager' : 'lazy'}">
          </div>
          <div class="produto-conteudo">
            <span class="produto-categoria">${categoria}</span>
            <h3>${produto}</h3>
            ${classificacao ? `<span class="produto-classificacao">Classificação ${classificacao}</span>` : ''}
            ${descricao ? `<p class="produto-descricao">${descricao}</p>` : ''}
            <span class="preco-original" data-preco="${preco}">R$ ${formatarPreco(preco)}</span>
            <a class="whatsapp" href="https://wa.me/5511999999999?text=${mensagem}" target="_blank" rel="noopener noreferrer">Comprar pelo WhatsApp</a>
          </div>
        </article>`;
    }).join('');

    return responder(e, { ok: true, html: html });
  } catch (erro) {
    return responder(e, {
      ok: false,
      erro: String(erro && erro.message ? erro.message : erro),
      html: '<div class="sem-produtos">Erro ao carregar os produtos.</div>'
    });
  }
}

function responder(e, payload) {
  const callback = e && e.parameter ? String(e.parameter.callback || '').trim() : '';

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return HtmlService
    .createHtmlOutput(payload.html || '')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ehPromocao(valor) {
  const normalizado = String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();

  return ['sim', 'true', 'verdadeiro', '1', 'yes', 'x'].includes(normalizado);
}

function normalizarPreco(valor) {
  let texto = String(valor || '').replace(/R\$/gi, '').trim();
  if (!texto) return 0;

  if (texto.includes(',')) {
    texto = texto.replace(/\./g, '').replace(',', '.');
  } else {
    texto = texto.replace(/[^0-9.]/g, '');
  }

  const numero = Number(texto);
  return Number.isFinite(numero) ? numero : 0;
}

function formatarPreco(valor) {
  return Number(valor || 0).toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function normalizarImagem(url) {
  const original = String(url || '').trim();
  if (!original) return 'IMG_7025.png';

  const match = original.match(/[-\w]{25,}/);
  if (original.includes('drive.google.com') && match) {
    return 'https://drive.google.com/thumbnail?id=' + match[0] + '&sz=w1200';
  }

  return escapeHtml(original);
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
