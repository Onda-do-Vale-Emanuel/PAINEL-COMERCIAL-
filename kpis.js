async function carregarJSON(caminho) {
  const resp = await fetch(caminho);
  return await resp.json();
}

function setTexto(id, valor) {
  const el = document.getElementById(id);
  if (el) el.innerText = valor;
}

async function atualizarKPIs() {
  try {
    const fat = await carregarJSON("site/dados/kpi_faturamento.json");
    const qtd = await carregarJSON("site/dados/kpi_quantidade_pedidos.json");

    setTexto("fatAtual", `R$ ${fat.atual.toLocaleString("pt-BR")}`);
    setTexto("fatAnoAnterior", `R$ ${fat.ano_anterior.toLocaleString("pt-BR")}`);
    setTexto("fatVariacao", `${fat.variacao}%`);
    setTexto("periodoAtual", `até ${fat.data_fim}`);
    setTexto("periodoAnterior", "até 28/01/2025");

    setTexto("qtdAtual", `${qtd.atual} pedidos`);
    setTexto("qtdAnoAnterior", `${qtd.ano_anterior} pedidos`);

    setTexto("qtdAtual2", qtd.atual);
    setTexto("qtdAnoAnterior2", qtd.ano_anterior);

    const meta = 1325000;
    const perc = ((fat.atual / meta) * 100).toFixed(1);
    setTexto("fatMetaPerc", `${perc}% da meta`);

    setTexto("dataAtualizacao", new Date().toLocaleString("pt-BR"));

  } catch (e) {
    console.error("Erro ao atualizar KPIs", e);
  }
}

atualizarKPIs();
