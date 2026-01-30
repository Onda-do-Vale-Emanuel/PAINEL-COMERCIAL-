// ======================================================
// FUNÇÕES AUXILIARES
// ======================================================
function carregarJSON(nome) {
  return fetch("site/dados/" + nome)
    .then((resp) => {
      if (!resp.ok) {
        throw new Error("Erro ao carregar " + nome);
      }
      return resp.json();
    })
    .catch((err) => {
      console.error(err);
      return null;
    });
}

function formatarMoeda(valor) {
  return valor.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 2,
  });
}

function formatarNumero(valor) {
  return valor.toLocaleString("pt-BR", {
    maximumFractionDigits: 0, // sem casas decimais — como você pediu no KG
  });
}

function formatarPercentual(valor) {
  return valor.toFixed(1).replace(".", ",") + "%";
}

function aplicarCorPosNeg(elemento, valor) {
  elemento.classList.remove("positivo", "negativo");
  if (valor > 0) elemento.classList.add("positivo");
  if (valor < 0) elemento.classList.add("negativo");
}

// ======================================================
// METAS POR MÊS (FATURAMENTO e KG)
// ======================================================
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
  12: { kg: 98000, fat: 1409516.02 },
};

// Função robusta para pegar a data do JSON
function obterDataReferencia(obj) {
  return (
    obj.data_atual || // caso 1
    obj.data ||       // caso 2
    obj.data_ref ||   // caso 3
    obj.periodo ||    // caso 4
    null
  );
}

// Extrai mês da data dd/mm/aaaa
function extrairMes(dataBR) {
  if (!dataBR || !dataBR.includes("/")) return 1;
  return Number(dataBR.split("/")[1]);
}

// ======================================================
// CARREGAR TODOS OS KPIs
// ======================================================
Promise.all([
  carregarJSON("kpi_faturamento.json"),
  carregarJSON("kpi_quantidade_pedidos.json"),
  carregarJSON("kpi_ticket_medio.json"),
  carregarJSON("kpi_kg_total.json"),
  carregarJSON("kpi_preco_medio.json"),
]).then(([fat, qtd, ticket, kg, preco]) => {

  if (!fat || !qtd || !ticket || !kg) {
    console.error("Algum JSON não foi carregado corretamente.");
    return;
  }

  // --------------------------------------------------
  // DATA DE REFERÊNCIA
  // --------------------------------------------------
  const dataRef = obterDataReferencia(fat);
  const mes = extrairMes(dataRef);
  const metasMes = METAS[mes];

  // --------------------------------------------------
  // SLIDE 1 – FATURAMENTO COM IPI
  // --------------------------------------------------
  const fatAtual = fat.atual;
  const fatAnterior = fat.ano_anterior;
  const fatVar = fat.variacao;

  const qtdAtual = qtd.atual;
  const qtdAnterior = qtd.ano_anterior;

  // Cards principais
  document.getElementById("fatQtdAtual").innerText =
    qtdAtual.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("fatValorAtual").innerText =
    formatarMoeda(fatAtual) + " (com IPI)";
  document.getElementById("fatDataAtual").innerText =
    "de 01/" + dataRef.substring(3) + " até " + dataRef;

  document.getElementById("fatQtdAnterior").innerText =
    qtdAnterior.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("fatValorAnterior").innerText =
    formatarMoeda(fatAnterior) + " (com IPI)";
  document.getElementById("fatDataAnterior").innerText =
    "de 01/" +
    fat.data_ano_anterior.substring(3) +
    " até " +
    fat.data_ano_anterior;

  // Variação
  const elFatVar = document.getElementById("fatVariacao");
  const prefixoFat = fatVar >= 0 ? "▲" : "▼";
  elFatVar.innerText =
    `${prefixoFat} ${formatarPercentual(Math.abs(fatVar))} vs ano anterior`;
  aplicarCorPosNeg(elFatVar, fatVar);

  // Meta FAT
  const metaFat = metasMes.fat;
  const percMetaFat = (fatAtual / metaFat) * 100;

  document.getElementById("fatMetaValor").innerText =
    "Meta mês: " + formatarMoeda(metaFat);

  const elFatMetaPerc = document.getElementById("fatMetaPerc");
  elFatMetaPerc.innerText =
    "🎯 " + formatarPercentual(percMetaFat) + " da meta";

  elFatMetaPerc.classList.remove("meta-ok", "meta-atencao", "meta-ruim");
  if (percMetaFat >= 100) elFatMetaPerc.classList.add("meta-ok");
  else if (percMetaFat >= 80) elFatMetaPerc.classList.add("meta-atencao");
  else elFatMetaPerc.classList.add("meta-ruim");

  // --------------------------------------------------
  // SLIDE 2 – KG TOTAL (com meta)
  // --------------------------------------------------
  const kgAtual = kg.atual;
  const kgAnterior = kg.ano_anterior;
  const kgVar = kg.variacao;

  const metaKg = metasMes.kg;
  const percMetaKg = (kgAtual / metaKg) * 100;

  document.getElementById("kgQtdAtual").innerText =
    qtdAtual.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("kgValorAtual").innerText =
    formatarNumero(kgAtual) + " kg";
  document.getElementById("kgDataAtual").innerText =
    "de 01/" + dataRef.substring(3) + " até " + dataRef;

  document.getElementById("kgQtdAnterior").innerText =
    qtdAnterior.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("kgValorAnterior").innerText =
    formatarNumero(kgAnterior) + " kg";
  document.getElementById("kgDataAnterior").innerText =
    "de 01/" +
    fat.data_ano_anterior.substring(3) +
    " até " +
    fat.data_ano_anterior;

  const elKgVar = document.getElementById("kgVariacao");
  const prefixoKg = kgVar >= 0 ? "▲" : "▼";
  elKgVar.innerText =
    `${prefixoKg} ${formatarPercentual(Math.abs(kgVar))} vs ano anterior`;
  aplicarCorPosNeg(elKgVar, kgVar);

  // meta KG
  document.getElementById("kgMetaValor").innerText =
    "Meta mês: " + formatarNumero(metaKg) + " kg";

  const elKgMetaPerc = document.getElementById("kgMetaPerc");
  elKgMetaPerc.innerText =
    "🎯 " + formatarPercentual(percMetaKg) + " da meta";

  elKgMetaPerc.classList.remove("meta-ok", "meta-atencao", "meta-ruim");
  if (percMetaKg >= 100) elKgMetaPerc.classList.add("meta-ok");
  else if (percMetaKg >= 80) elKgMetaPerc.classList.add("meta-atencao");
  else elKgMetaPerc.classList.add("meta-ruim");

  // --------------------------------------------------
  // SLIDE 3 – TICKET MÉDIO
  // --------------------------------------------------
  const ticketAtual = ticket.atual;
  const ticketAnterior = ticket.ano_anterior;
  const ticketVar = ticket.variacao;

  document.getElementById("ticketAtual").innerText =
    formatarMoeda(ticketAtual);
  document.getElementById("ticketAnterior").innerText =
    formatarMoeda(ticketAnterior);

  document.getElementById("ticketQtdAtual").innerText =
    qtdAtual.toLocaleString("pt-BR") + " pedidos no período";
  document.getElementById("ticketQtdAnterior").innerText =
    qtdAnterior.toLocaleString("pt-BR") + " pedidos no período";

  const elTicketVar = document.getElementById("ticketVariacao");
  const prefixoTicket = ticketVar >= 0 ? "▲" : "▼";
  elTicketVar.innerText =
    `${prefixoTicket} ${formatarPercentual(Math.abs(ticketVar))} vs ano anterior`;
  aplicarCorPosNeg(elTicketVar, ticketVar);


  // --------------------------------------------------
  // SLIDE 4 – PREÇO MÉDIO (NOVO)
  // --------------------------------------------------
  if (preco) {
    document.getElementById("preco-medio-kg").innerText =
      "R$ " + preco.preco_medio_kg.toLocaleString("pt-BR");

    document.getElementById("preco-medio-m2").innerText =
      "R$ " + preco.preco_medio_m2.toLocaleString("pt-BR");
  }
});
