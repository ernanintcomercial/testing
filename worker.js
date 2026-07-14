// worker.js
//
// Coloque isto em Workers & Pages > Create > Worker (cole direto no editor
// "Quick edit" do painel da Cloudflare, sem precisar instalar nada no
// computador). Não usa nenhum recurso pago — o plano Free do Workers não
// pede cartão de crédito.
//
// O que ele faz: antes de mostrar o site, pede usuário e senha (autenticação
// HTTP Basic, suportada nativamente por qualquer navegador). Só depois de
// digitar certo é que ele repassa a requisição pro GitHub Pages de verdade.

export default {
  async fetch(request, env) {
    const cabecalhoAuth = request.headers.get("Authorization");

    if (!autenticado(cabecalhoAuth, env.AUTH_USER, env.AUTH_PASS)) {
      return new Response("Acesso restrito.", {
        status: 401,
        headers: {
          "WWW-Authenticate": 'Basic realm="Painel de Vendas"',
        },
      });
    }

    // Repassa a requisição pro GitHub Pages, mantendo o mesmo caminho
    // (ex.: /resumo.json, /index.html) e devolve a resposta como se fosse
    // deste próprio domínio.
    const url = new URL(request.url);
    const destino = `https://${env.GITHUB_PAGES_HOST}${url.pathname}${url.search}`;
    const resposta = await fetch(destino, {
      headers: request.headers,
    });

    // Copia a resposta pra poder devolver (o corpo original só pode ser
    // lido uma vez).
    return new Response(resposta.body, {
      status: resposta.status,
      headers: resposta.headers,
    });
  },
};

function autenticado(cabecalho, usuarioEsperado, senhaEsperada) {
  if (!cabecalho || !cabecalho.startsWith("Basic ")) return false;
  try {
    const decodificado = atob(cabecalho.slice(6));
    const separador = decodificado.indexOf(":");
    const usuario = decodificado.slice(0, separador);
    const senha = decodificado.slice(separador + 1);
    return usuario === usuarioEsperado && senha === senhaEsperada;
  } catch (e) {
    return false;
  }
}
