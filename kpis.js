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
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
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
// METAS POR MÊS
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

function obterMetaMes(dataBR) {
  if (!dataBR) return METAS[1];
  const mes = Number(dataBR.split("/")[1]);
  return METAS[mes] || METAS[1];
}

// ======================================================
// CARREGA TODOS OS ARQUIVOS JSON
// ======================================================
Promise.all([
  carregarJSON("kpi_faturamento.json"),
  carregarJSON("kpi_quantidade_pedidos.json"),
  carregarJSON("kpi_ticket_medio.json"),
  carregarJSON("kpi_kg_total.json"),
  carregarJSON("kpi_preco_medio.json")  // ⭐ AGORA TÁ AQUI
]).then(([fat, qtd, ticket, kg, preco]) => {

  if (!fat || !qtd || !ticket || !kg) {
    console.error("Erro: algum JSON principal não carregou.");
    return;
  }

  // --------------------------------------------------
  // SLIDE 1 — FATURAMENTO
  // --------------------------------------------------
  const dataRef = fat.data_atual;
  const metasMes = obterMetaMes(dataRef);

  document.getElementById("fatQtdAtual").innerText =
    qtd.atual.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("fatValorAtual").innerText =
    formatarMoeda(fat.atual) + " (com IPI)";
  document.getElementById("fatDataAtual").innerText =
    `de 01/${dataRef.substring(3)} até ${dataRef}`;

  document.getElementById("fatQtdAnterior").innerText =
    qtd.ano_anterior.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("fatValorAnterior").innerText =
    formatarMoeda(fat.ano_anterior) + " (com IPI)";
  document.getElementById("fatDataAnterior").innerText =
    `de 01/${fat.data_ano_anterior.substring(3)} até ${fat.data_ano_anterior}`;

  const elFatVar = document.getElementById("fatVariacao");
  const prefixoFat = fat.variacao >= 0 ? "▲" : "▼";
  elFatVar.innerText = `${prefixoFat} ${formatarPercentual(Math.abs(fat.variacao))} vs ano anterior`;
  aplicarCorPosNeg(elFatVar, fat.variacao);

  const metaFat = metasMes.fat;
  const percMetaFat = (fat.atual / metaFat) * 100;

  document.getElementById("fatMetaValor").innerText =
    "Meta mês: " + formatarMoeda(metaFat);
  const elFatMetaPerc = document.getElementById("fatMetaPerc");
  elFatMetaPerc.innerText =
    "🎯 " + formatarPercentual(percMetaFat) + " da meta";

  // --------------------------------------------------
  // SLIDE 2 — KG TOTAL
  // --------------------------------------------------
  document.getElementById("kgQtdAtual").innerText =
    qtd.atual.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("kgValorAtual").innerText =
    formatarNumero(kg.atual) + " kg";
  document.getElementById("kgDataAtual").innerText =
    `de 01/${dataRef.substring(3)} até ${dataRef}`;

  document.getElementById("kgQtdAnterior").innerText =
    qtd.ano_anterior.toLocaleString("pt-BR") + " pedidos";
  document.getElementById("kgValorAnterior").innerText =
    formatarNumero(kg.ano_anterior) + " kg";
  document.getElementById("kgDataAnterior").innerText =
    `de 01/${fat.data_ano_anterior.substring(3)} até ${fat.data_ano_anterior}`;

  const elKgVar = document.getElementById("kgVariacao");
  const prefixoKg = kg.variacao >= 0 ? "▲" : "▼";
  elKgVar.innerText = `${prefixoKg} ${formatarPercentual(Math.abs(kg.variacao))} vs ano anterior`;
  aplicarCorPosNeg(elKgVar, kg.variacao);

  const metaKg = metasMes.kg;
  const percMetaKg = (kg.atual / metaKg) * 100;

  document.getElementById("kgMetaValor").innerText =
    "Meta mês: " + formatarNumero(metaKg) + " kg";

  const elKgMetaPerc = document.getElementById("kgMetaPerc");
  elKgMetaPerc.innerText =
    "🎯 " + formatarPercentual(percMetaKg) + " da meta";

  // --------------------------------------------------
  // SLIDE 3 — TICKET MÉDIO
  // --------------------------------------------------
  document.getElementById("ticketAtual").innerText =
    formatarMoeda(ticket.atual);
  document.getElementById("ticketAnterior").innerText =
    formatarMoeda(ticket.ano_anterior);

  const elTicketVar = document.getElementById("ticketVariacao");
  const prefixoTicket = ticket.variacao >= 0 ? "▲" : "▼";
  elTicketVar.innerText =
    `${prefixoTicket} ${formatarPercentual(Math.abs(ticket.variacao))} vs ano anterior`;
  aplicarCorPosNeg(elTicketVar, ticket.variacao);

  // --------------------------------------------------
  // ⭐ SLIDE 4 — PREÇO MÉDIO (NOVO)
  // --------------------------------------------------
  if (preco) {
    document.getElementById("precoKg").innerText =
      "R$ " + preco.preco_medio_kg.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
    document.getElementById("precoM2").innerText =
      "R$ " + preco.preco_medio_m2.toLocaleString("pt-BR", { minimumFractionDigits: 2 });
  }
});
