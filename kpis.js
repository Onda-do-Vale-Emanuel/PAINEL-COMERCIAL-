document.addEventListener("DOMContentLoaded", () => {
  atualizarKPIs();
});

async function atualizarKPIs() {
  try {
    const fat = await fetch("site/dados/kpi_faturamento.json").then(r => r.json());
    const qtd = await fetch("site/dados/kpi_quantidade_pedidos.json").then(r => r.json());

    // FATURAMENTO
    document.getElementById("fatAtual").innerText =
      fat.atual.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

    document.getElementById("fatAnoAnterior").innerText =
      fat.ano_anterior.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

    document.getElementById("fatVariacao").innerText =
      fat.variacao !== null ? `▲ ${fat.variacao}% vs ano anterior` : "--";

    document.getElementById("periodoAtual").innerText =
      `até ${fat.data_fim}`;

    document.getElementById("periodoAnterior").innerText =
      `até ${fat.data_fim.replace("2026", "2025")}`;

    // QUANTIDADE
    document.getElementById("qtdAtual").innerText =
      `${qtd.atual} pedidos`;

    document.getElementById("qtdAnoAnterior").innerText =
      `${qtd.ano_anterior} pedidos`;

  } catch (erro) {
    console.error("Erro ao atualizar KPIs:", erro);
  }
}
