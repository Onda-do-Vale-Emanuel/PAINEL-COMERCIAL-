async function carregarKPIs() {
    async function loadJSON(path) {
        try {
            const r = await fetch(path);
            if (!r.ok) return null;
            return await r.json();
        } catch (e) {
            console.error("Erro ao carregar JSON:", path, e);
            return null;
        }
    }

    return {
        pedidos: await loadJSON("dados/kpi_quantidade_pedidos.json"),
        ticket: await loadJSON("dados/kpi_ticket_medio.json"),
        kg: await loadJSON("dados/kpi_kg_total.json"),
        preco: await loadJSON("dados/kpi_preco_medio.json")  // NOVO
    };
}

function atualizarTela(d) {
    // --- SLIDE 1 ---
    document.getElementById("kpi-pedidos").innerText =
        d.pedidos?.atual?.toLocaleString("pt-BR") ?? "--";
    document.getElementById("var-pedidos").innerText =
        d.pedidos?.variacao ? `${d.pedidos.variacao}%` : "--";

    document.getElementById("kpi-ticket").innerText =
        d.ticket?.atual?.toLocaleString("pt-BR") ?? "--";
    document.getElementById("var-ticket").innerText =
        d.ticket?.variacao ? `${d.ticket.variacao}%` : "--";

    document.getElementById("kpi-kg").innerText =
        d.kg?.atual?.toLocaleString("pt-BR") ?? "--";
    document.getElementById("var-kg").innerText =
        d.kg?.variacao ? `${d.kg.variacao}%` : "--";

    // --- SLIDE PREÇO MÉDIO ---
    if (d.preco) {
        document.getElementById("preco-medio-kg").innerText =
            `R$ ${d.preco.preco_medio_kg.toLocaleString("pt-BR")}`;

        document.getElementById("preco-medio-m2").innerText =
            `R$ ${d.preco.preco_medio_m2.toLocaleString("pt-BR")}`;
    }
}

(async () => {
    const dados = await carregarKPIs();
    atualizarTela(dados);
})();
