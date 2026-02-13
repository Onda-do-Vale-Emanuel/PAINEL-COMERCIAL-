function carregarJSON(nome) {
  return fetch("site/dados/" + nome)
    .then(r => r.json())
    .catch(() => null);
}

function formatarMoeda(v) {
  return Number(v).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL"
  });
}

function formatarNumero(v) {
  return Number(v).toLocaleString("pt-BR", {
    maximumFractionDigits: 0
  });
}

function formatarPercentual(v) {
  return Number(v).toFixed(1).replace(".", ",") + "%";
}

function aplicarCorPosNeg(el, valor) {
  el.classList.remove("positivo", "negativo");
  if (valor > 0) el.classList.add("positivo");
  if (valor < 0) el.classList.add("negativo");
}

Promise.all([
  carregarJSON("kpi_faturamento.json"),
  carregarJSON("kpi_quantidade_pedidos.json"),
  carregarJSON("kpi_ticket_medio.json"),
  carregarJSON("kpi_kg_total.json"),
  carregarJSON("kpi_preco_medio.json")
]).then(([fat, qtd, ticket, kg, preco]) => {

  if (!fat || !qtd || !ticket || !kg) return;

  /* ================= METAS ================= */
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
    10:{ kg: 150000, fat: 2199335.81 },
    11:{ kg: 150000, fat: 2199360.62 },
    12:{ kg: 98000,  fat: 1409516.02 }
  };

  const mes = Number(fat.data_atual.substring(3,5));

  /* ================= SLIDE 1 ================= */
  document.getElementById("fatQtdAtual").innerText = qtd.atual + " pedidos";
  document.getElementById("fatValorAtual").innerText = formatarMoeda(fat.atual);
  document.getElementById("fatDataAtual").innerText =
    "de " + fat.inicio_mes + " até " + fat.data_atual;

  document.getElementById("fatQtdAnterior").innerText =
    qtd.ano_anterior + " pedidos";
  document.getElementById("fatValorAnterior").innerText =
    formatarMoeda(fat.ano_anterior);
  document.getElementById("fatDataAnterior").innerText =
    "de " + fat.inicio_mes_anterior + " até " + fat.data_ano_anterior;

  const elFatVar = document.getElementById("fatVariacao");
  const pfFat = fat.variacao >= 0 ? "▲" : "▼";
  elFatVar.innerText =
    `${pfFat} ${formatarPercentual(Math.abs(fat.variacao))} vs ano anterior`;
  aplicarCorPosNeg(elFatVar, fat.variacao);

  /* META FAT */
  const metaFat = METAS[mes].fat;
  const metaFatPerc = (fat.atual / metaFat) * 100;

  document.getElementById("fatMetaValor").innerText =
    "Meta mês: " + formatarMoeda(metaFat);

  const elFatMetaPerc = document.getElementById("fatMetaPerc");
  elFatMetaPerc.innerText =
    "🎯 " + metaFatPerc.toFixed(1).replace(".", ",") + "% da meta";

  /* ================= SLIDE 2 ================= */
  document.getElementById("kgQtdAtual").innerText = qtd.atual + " pedidos";
  document.getElementById("kgValorAtual").innerText =
    formatarNumero(kg.atual) + " kg";
  document.getElementById("kgDataAtual").innerText =
    "de " + fat.inicio_mes + " até " + fat.data_atual;

  document.getElementById("kgQtdAnterior").innerText =
    qtd.ano_anterior + " pedidos";
  document.getElementById("kgValorAnterior").innerText =
    formatarNumero(kg.ano_anterior) + " kg";
  document.getElementById("kgDataAnterior").innerText =
    "de " + fat.inicio_mes_anterior + " até " + fat.data_ano_anterior;

  const elKgVar = document.getElementById("kgVariacao");
  const pfKG = kg.variacao >= 0 ? "▲" : "▼";
  elKgVar.innerText =
    `${pfKG} ${formatarPercentual(Math.abs(kg.variacao))} vs ano anterior`;
  aplicarCorPosNeg(elKgVar, kg.variacao);

  /* META KG */
  const metaKG = METAS[mes].kg;
  const metaKGperc = (kg.atual / metaKG) * 100;

  document.getElementById("kgMetaValor").innerText =
    "Meta mês: " + formatarNumero(metaKG) + " kg";

  const elKgMetaPerc = document.getElementById("kgMetaPerc");
  elKgMetaPerc.innerText =
    "🎯 " + metaKGperc.toFixed(1).replace(".", ",") + "% da meta";

  /* ================= SLIDE 3 ================= */
  document.getElementById("ticketAtual").innerText =
    formatarMoeda(ticket.atual);
  document.getElementById("ticketAnterior").innerText =
    formatarMoeda(ticket.ano_anterior);

  document.getElementById("ticketQtdAtual").innerText =
    qtd.atual + " pedidos no período";
  document.getElementById("ticketQtdAnterior").innerText =
    qtd.ano_anterior + " pedidos no período";

  const elTicketVar = document.getElementById("ticketVariacao");
  const pfT = ticket.variacao >= 0 ? "▲" : "▼";
  elTicketVar.innerText =
    `${pfT} ${formatarPercentual(Math.abs(ticket.variacao))} vs ano anterior`;
  aplicarCorPosNeg(elTicketVar, ticket.variacao);

  /* ================= SLIDE 4 ================= */
  if (preco && preco.atual) {

    document.getElementById("precoMedioKG").innerText =
      formatarMoeda(preco.atual.preco_medio_kg);

    document.getElementById("precoMedioM2").innerText =
      formatarMoeda(preco.atual.preco_medio_m2);

    document.getElementById("precoDataKG").innerText =
      "Atualizado até " + preco.atual.data;

    document.getElementById("precoDataM2").innerText =
      "Atualizado até " + preco.atual.data;
  }

  if (preco && preco.ano_anterior) {

    document.getElementById("precoMedioKGant").innerText =
      formatarMoeda(preco.ano_anterior.preco_medio_kg);

    document.getElementById("precoMedioM2ant").innerText =
      formatarMoeda(preco.ano_anterior.preco_medio_m2);

    document.getElementById("precoDataKGant").innerText =
      "Atualizado até " + preco.ano_anterior.data;

    document.getElementById("precoDataM2ant").innerText =
      "Atualizado até " + preco.ano_anterior.data;
  }

});
