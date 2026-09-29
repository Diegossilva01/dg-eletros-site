# DG Eletros — GitHub + Vercel + Neon

O site e o painel `/admin.html` são servidos pela Vercel. As rotas `/api/*` consultam o PostgreSQL do Neon. O GitHub guarda somente o código. Dados de clientes, usuários, vendas e fotos ficam no Neon.

## Preparação

1. Faça um backup da planilha e do banco Neon. Suspenda cadastros e vendas durante o corte para não perder alterações feitas depois da exportação.
2. No **Editor SQL do projeto Neon**, execute `schema.sql`.
3. Em seguida, execute o arquivo **privado** `importacao-dg-eletros-privada.sql` fornecido separadamente. Ele contém informações pessoais e hashes de senha; nunca faça commit dele. As seis contagens esperadas aparecem no fim do arquivo. O SQL usa `ON CONFLICT DO NOTHING`, então não duplica registros se for executado novamente.
4. Na Vercel, importe ou conecte o repositório GitHub da DG Eletros. O diretório raiz deve conter `index.html`, `package.json` e `api/`. Configure o projeto como **Other** e deixe os comandos de build/output vazios.
5. Em **Settings → Environment Variables**, adicione `DATABASE_URL` como Secret para **Production**, com a string de conexão do projeto Neon que recebeu o SQL. A senha fica somente na Vercel. Faça um novo deploy.
6. Abra `/admin.html`, entre com um usuário migrado e confirme produtos, vendas, contas, comissões, horas extras e o catálogo público. Faça uma venda fictícia apenas em uma cópia de teste do banco e verifique estoque e recebimento antes do corte real.

## Fotos antigas

A planilha possui links de imagens do Drive, mas não os arquivos. Após importar, execute **uma vez** `node scripts/migrar-fotos.js` em um computador com Node.js, `npm ci` e a variável `DATABASE_URL` definida no ambiente. O script tenta copiar cada imagem pública para `dge_photos` no Neon e substitui o link por `/api/foto?id=...`. Ele informa quais imagens estão privadas, inválidas ou grandes demais. Essas devem ser baixadas manualmente e reenviadas pelo painel antes de desativar o Drive. Não cole a conexão do banco em chat, código ou commit.

Novas fotos enviadas pelo painel ficam no Neon. Cada arquivo é limitado a 500 KB e cada cadastro aceita até seis fotos. A Vercel limita o corpo de uma requisição a 4,5 MB. Fotos antigas não migradas ainda dependem do Drive.

## Conferência do corte

- No Neon, confirme `Produtos=115`, `Vendas=146`, `Usuarios=4`, `LogExclusoes=76`, `Pagamentos Comissoes=1`, `Horas Extras=1`, relativos ao arquivo recebido. Uma exportação mais nova pode ter números diferentes.
- Confira os valores e saldos com a planilha, produtos ocultos/reservados e permissões de cada perfil.
- Confira que todos os links de `Fotos` começaram com `/api/foto` ou foram corrigidos manualmente. Produtos sem fotos na origem continuarão sem fotos.
- Só depois suspenda Apps Script/planilha como fonte de dados. Guarde a planilha original como backup privado.

Este pacote prepara a migração, mas o banco e o domínio em produção só mudam quando o SQL privado é executado e o repositório é implantado na sua conta Vercel.
