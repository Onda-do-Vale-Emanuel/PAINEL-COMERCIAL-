function carregarJSON(nome) {
  return fetch("site/dados/" + nome)
    .then((resp) => resp.json())
    .catch(() => null);
}

function moeda(v) {
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function numero(v) {
  return v.toLocaleString("pt-BR", { maximumFractionDigits: 0 });
}

function percentual(v) {
  return v.toFixed(1).replace(".", ",") + "%";
}

// APLICA COR NO TRIÂNGULO + TEXTO
function corVariacao(elemento, valor) {
  elemento.classList.remove("positivo", "negativo");
  if (valor > 0) elemento.classList.add("positivo");
  if (valor < 0) elemento.classList.add("negativo");
}

const METAS = {
  1: { kg: 100000, fat: 1324746.56 },
  2: { kg: 100000, fat: 1324746.56 },
  3: { kg: 120000, fat: 1598757.69 },
  4: { kg: 130000, fat: 1910459.23 },
  5: { kg: 130000, fat: 1892998.21 },
  6: { kg: 130000, fat: 1892995.74 },
  7: { kg: 150000, fat: 2199365.46 },
  8: { kg: 150000, fat: 2199350.47 },
  9: { kg: 150000, fat: 2199340.46 },
  10: { kg: 150000, fat: 2199335.81 },
  11: { kg: 150000, fat: 2199360.62 },
  12: { kg: 98000, fat: 1409516.02 }
};

Promise.all([
  carregarJSON("kpi_faturamento.json"),
  carregarJSON("kpi_quantidade_pedidos.json"),
  carregarJSON("kpi_ticket_medio.json"),
  carregarJSON("kpi_kg_total.json"),
  carregarJSON("kpi_preco_medio.json"),
]).then(([fat, qtd, ticket, kg, preco]) => {

  const mes = Number(fat.data_atual.split("/")[1]);
  const meta = METAS[mes];

  // ============ FATURAMENTO ============
  const fatVar = fat.variacao;
  const elFatVar = document.getElementById("fatVariacao");
  elFatVar.innerText = `${fatVar >= 0 ? "▲" : "▼"} ${percentual(Math.abs(fatVar))} vs ano anterior`;
  corVariacao(elFatVar, fatVar);

  // ============ KG ============
  const kgVar = kg.variacao;
  const elKgVar = document.getElementById("kgVariacao");
  elKgVar.innerText = `${kgVar >= 0 ? "▲" : "▼"} ${percentual(Math.abs(kgVar))} vs ano anterior`;
  corVariacao(elKgVar, kgVar);

  // ============ TICKET ============
  const ticketVar = ticket.variacao;
  const elTicketVar = document.getElementById("ticketVariacao");
  elTicketVar.innerText = `${ticketVar >= 0 ? "▲" : "▼"} ${percentual(Math.abs(ticketVar))} vs ano anterior`;
  corVariacao(elTicketVar, ticketVar);

  // DEMAIS CAMPOS (iguais antes)
  document.getElementById("fatQtdAtual").innerText = qtd.atual + " pedidos";
  document.getElementById("fatValorAtual").innerText = moeda(fat.atual);
  document.getElementById("fatDataAtual").innerText =
    `de 01/${fat.data_atual.substring(3)} até ${fat.data_atual}`;

  document.getElementById("fatQtdAnterior").innerText = qtd.ano_anterior + " pedidos";
  document.getElementById("fatValorAnterior").innerText = moeda(fat.ano_anterior);
  document.getElementById("fatDataAnterior").innerText =
    `de 01/${fat.data_ano_anterior.substring(3)} até ${fat.data_ano_anterior}`;

  document.getElementById("kgQtdAtual").innerText = qtd.atual + " pedidos";
  document.getElementById("kgValorAtual").innerText = numero(kg.atual) + " kg";

  document.getElementById("kgQtdAnterior").innerText = qtd.ano_anterior + " pedidos";
  document.getElementById("kgValorAnterior").innerText = numero(kg.ano_anterior) + " kg";

  document.getElementById("ticketAtual").innerText = moeda(ticket.atual);
  document.getElementById("ticketAnterior").innerText = moeda(ticket.ano_anterior);

  if (preco) {
    document.getElementById("precoKg").innerText = moeda(preco.preco_medio_kg);
    document.getElementById("precoKgInfo").innerText =
      numero(preco.total_kg) + " kg no período";

    document.getElementById("precoM2").innerText = moeda(preco.preco_medio_m2);
    document.getElementById("precoM2Info").innerText =
      numero(preco.total_m2) + " m² no período";
  }
});
