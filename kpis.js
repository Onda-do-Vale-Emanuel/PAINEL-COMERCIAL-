function carregarJSON(nome) {
  return fetch("site/dados/" + nome).then(r => r.json());
}

function moeda(v) {
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function num(v) {
  return v.toLocaleString("pt-BR");
}

Promise.all([
  carregarJSON("kpi_faturamento.json"),
  carregarJSON("kpi_quantidade_pedidos.json"),
  carregarJSON("kpi_kg_total.json"),
  carregarJSON("kpi_preco_medio.json")
]).then(([fat, qtd, kg, preco]) => {

  document.getElementById("fatQtdAtual").innerText = qtd.atual + " pedidos";
  document.getElementById("fatValorAtual").innerText = moeda(fat.atual);
  document.getElementById("fatDataAtual").innerText =
    `de ${fat.inicio_mes} até ${fat.data_atual}`;

  document.getElementById("fatQtdAnterior").innerText = qtd.ano_anterior + " pedidos";
  document.getElementById("fatValorAnterior").innerText = moeda(fat.ano_anterior);
  document.getElementById("fatDataAnterior").innerText =
    `de ${fat.inicio_mes_anterior} até ${fat.data_ano_anterior}`;

  document.getElementById("kgQtdAtual").innerText = qtd.atual + " pedidos";
  document.getElementById("kgValorAtual").innerText = num(kg.atual) + " kg";
  document.getElementById("kgDataAtual").innerText =
    `de ${fat.inicio_mes} até ${fat.data_atual}`;

  document.getElementById("kgQtdAnterior").innerText = qtd.ano_anterior + " pedidos";
  document.getElementById("kgValorAnterior").innerText = num(kg.ano_anterior) + " kg";
  document.getElementById("kgDataAnterior").innerText =
    `de ${fat.inicio_mes_anterior} até ${fat.data_ano_anterior}`;

  document.getElementById("precoMedioKG").innerText =
    "R$ " + preco.preco_medio_kg.toLocaleString("pt-BR");

  document.getElementById("precoMedioM2").innerText =
    "R$ " + preco.preco_medio_m2.toLocaleString("pt-BR");

  document.getElementById("precoDataKG").innerText = "Atualizado até " + preco.data;
  document.getElementById("precoDataM2").innerText = "Atualizado até " + preco.data;
});
