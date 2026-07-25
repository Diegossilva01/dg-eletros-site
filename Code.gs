function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Produtos') || ss.getSheets()[0];
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return HtmlService.createHtmlOutput('<div class="sem-produtos">Nenhum produto cadastrado.</div>');
  }

  const rows = values.slice(1).filter(row => String(row[0] || '').trim());

  const html = rows.map((row, index) => {
    const produto = escapeHtml(row[0]);
    const categoria = escapeHtml(row[1]);
    const precoTexto = String(row[2] || '').trim();
    const classificacao = escapeHtml(row[3]);
    const descricao = escapeHtml(row[4]);
    const imagem = normalizarImagem(row[5]);
    const promocaoOriginal = String(row[6] || '').trim();
    const promocao = ehPromocao(promocaoOriginal) ? 'sim' : 'nao';
    const preco = normalizarPreco(precoTexto);

    const mensagem = encodeURIComponent(
      'Olá! Tenho interesse no produto: ' + String(row[0] || '').trim() +
      (classificacao ? ' - Classificação ' + String(row[3] || '').trim() : '')
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

  return HtmlService.createHtmlOutput(html)
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
    return 'https://drive.google.com/uc?export=view&id=' + match[0];
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
