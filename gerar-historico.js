// gerar-historico.js
//
// Roda depois de cada atualização de dados.xlsx / INDEX.xlsx /
// Clientes_Prioritarios.xlsx. Recalcula os mesmos números que o site mostra
// hoje (Pedidos em Aberto R$, Qtd. de pedidos, Clientes prioritários
// afetados, Dias em aberto máx. — no total e por região) e acrescenta um
// snapshot do dia dentro de historico.json.
//
// IMPORTANTE: a lógica de filtro/agrupamento aqui precisa ficar igual à do
// index.html. Se um dia mudar uma regra lá (ex.: DIAS_MIN), replique aqui
// também, senão o histórico e o "hoje" da tela passam a divergir.

const fs = require('fs');
const XLSX = require('xlsx');

const DIAS_MIN = 25;
const REGION_LABELS = {
    'SUDESTE': 'SUDESTE',
    'SUL CENTRO OESTE': 'SUL E CENTRO OESTE',
    'NORTE NORDESTE': 'NORTE E NORDESTE',
};
const REGION_FALLBACK = 'SEM REGIÃO / NÃO MAPEADO';
const REGION_ORDER = ['SUDESTE', 'SUL E CENTRO OESTE', 'NORTE E NORDESTE'];
const UF_TO_REGION = {
    'ESPIRITO SANTO': 'SUDESTE', 'MINAS GERAIS': 'SUDESTE', 'RIO DE JANEIRO': 'SUDESTE', 'SAO PAULO': 'SUDESTE',
    'PARANA': 'SUL E CENTRO OESTE', 'RIO GRANDE DO SUL': 'SUL E CENTRO OESTE', 'SANTA CATARINA': 'SUL E CENTRO OESTE',
    'DISTRITO FEDERAL': 'SUL E CENTRO OESTE', 'GOIAS': 'SUL E CENTRO OESTE', 'MATO GROSSO': 'SUL E CENTRO OESTE', 'MATO GROSSO DO SUL': 'SUL E CENTRO OESTE',
    'ACRE': 'NORTE E NORDESTE', 'AMAPA': 'NORTE E NORDESTE', 'AMAZONAS': 'NORTE E NORDESTE', 'PARA': 'NORTE E NORDESTE',
    'RONDONIA': 'NORTE E NORDESTE', 'RORAIMA': 'NORTE E NORDESTE', 'TOCANTINS': 'NORTE E NORDESTE', 'TOCANTIS': 'NORTE E NORDESTE',
    'ALAGOAS': 'NORTE E NORDESTE', 'BAHIA': 'NORTE E NORDESTE', 'CEARA': 'NORTE E NORDESTE', 'MARANHAO': 'NORTE E NORDESTE',
    'PARAIBA': 'NORTE E NORDESTE', 'PERNAMBUCO': 'NORTE E NORDESTE', 'PIAUI': 'NORTE E NORDESTE', 'RIO GRANDE DO NORTE': 'NORTE E NORDESTE', 'SERGIPE': 'NORTE E NORDESTE',
};

const HISTORICO_ARQUIVO = 'historico.json';
const HISTORICO_MAX_DIAS = 400; // ~13 meses de retenção, evita crescer pra sempre

function normalizeUF(raw) {
    if (!raw) return null;
    return String(raw).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
}
function regionFromUF(raw) {
    const norm = normalizeUF(raw);
    return norm ? (UF_TO_REGION[norm] || null) : null;
}
function normalizeRegion(raw) {
    if (!raw) return null;
    const norm = String(raw).trim().replace(/\s+/g, ' ');
    return REGION_LABELS[norm] || norm;
}
function leadingCode(str) {
    if (str === null || str === undefined) return null;
    const m = String(str).trim().match(/^(\d+)/);
    return m ? parseInt(m[1], 10) : null;
}
function toUTCDay(d) {
    if (!(d instanceof Date) || isNaN(d)) return null;
    return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
function excelSerialToDate(value) {
    if (value === null || value === undefined || value === '') return null;
    if (value instanceof Date) return isNaN(value) ? null : value;
    const n = typeof value === 'number' ? value : parseFloat(value);
    if (isNaN(n)) return null;
    const days = Math.floor(n);
    return new Date(Date.UTC(1899, 11, 30) + days * 86400000);
}

function sheetRows(wb, sheetName) {
    const name = sheetName || wb.SheetNames[0];
    const ws = wb.Sheets[name];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

function main() {
    const wbDados = XLSX.readFile('dados.xlsx');
    const wbIndex = XLSX.readFile('INDEX.xlsx');
    const wbClientes = XLSX.readFile('Clientes_Prioritarios.xlsx');

    const idxRows = sheetRows(wbIndex);
    const repRegionMap = new Map();
    for (let i = 1; i < idxRows.length; i++) {
        const row = idxRows[i];
        if (!row || row[0] === null) continue;
        const codigo = parseInt(row[0], 10);
        if (isNaN(codigo)) continue;
        repRegionMap.set(codigo, normalizeRegion(row[3]));
    }

    const cliRows = sheetRows(wbClientes);
    const prioritySet = new Set();
    for (let i = 1; i < cliRows.length; i++) {
        const row = cliRows[i];
        if (!row || row[1] === null) continue;
        const cod = parseInt(row[1], 10);
        if (!isNaN(cod)) prioritySet.add(cod);
    }

    const raw = sheetRows(wbDados, 'Export') || sheetRows(wbDados);
    const now = new Date();
    const hojeUTC = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate());

    let totalValor = 0;
    let maxDiasGlobal = 0;
    const clientesUnicos = new Set();
    const pedidosUnicos = new Set(); // chave composta pra contar "qtd de pedidos" igual à árvore do site
    const porRegiao = new Map(); // label -> { valor, pedidosSet, diasMax }

    for (let i = 1; i < raw.length; i++) {
        const row = raw[i];
        if (!row) continue;
        const nrPedido = row[4];
        const dtImplantRaw = row[1];
        if (nrPedido === null || nrPedido === '' || isNaN(parseInt(nrPedido, 10))) continue;
        const dtImplant = excelSerialToDate(dtImplantRaw);
        if (!(dtImplant instanceof Date) || isNaN(dtImplant)) continue;

        const dtImplantUTC = toUTCDay(dtImplant);
        const diasAberto = Math.round((hojeUTC - dtImplantUTC) / 86400000);
        if (diasAberto < DIAS_MIN) continue;

        const qtdAberta = parseFloat(row[29]) || 0;
        if (qtdAberta <= 0) continue;

        const grupoCode = leadingCode(row[25]);
        if (grupoCode === null || !prioritySet.has(grupoCode)) continue;

        const repCode = leadingCode(row[21]);
        const repLabel = row[21] !== null ? String(row[21]).trim() : '(sem representante)';
        let regionLabel;
        if (repCode !== null && repRegionMap.has(repCode)) {
            regionLabel = repRegionMap.get(repCode);
        } else {
            regionLabel = regionFromUF(row[19]) || REGION_FALLBACK;
        }

        const clienteLabel = row[17] !== null ? String(row[17]).trim() : '(sem cliente)';
        const contratoLabel = row[23] !== null ? String(row[23]).trim() : '(sem contrato)';
        const pedidoLabel = row[5] !== null ? String(row[5]).trim() : '(sem nº de pedido)';
        const precoLiquido = parseFloat(row[33]) || 0;
        const valorItem = qtdAberta * precoLiquido;

        totalValor += valorItem;
        if (diasAberto > maxDiasGlobal) maxDiasGlobal = diasAberto;
        clientesUnicos.add(clienteLabel);

        const chavePedido = [regionLabel, repLabel, clienteLabel, contratoLabel, pedidoLabel].join('||');
        pedidosUnicos.add(chavePedido);

        if (!porRegiao.has(regionLabel)) porRegiao.set(regionLabel, { valor: 0, pedidosSet: new Set(), diasMax: 0 });
        const r = porRegiao.get(regionLabel);
        r.valor += valorItem;
        r.pedidosSet.add(chavePedido);
        if (diasAberto > r.diasMax) r.diasMax = diasAberto;
    }

    const regioesOrdenadas = [
        ...REGION_ORDER.filter(l => porRegiao.has(l)),
        ...[...porRegiao.keys()].filter(l => !REGION_ORDER.includes(l)),
    ];

    const snapshot = {
        data: new Date().toISOString().slice(0, 10), // AAAA-MM-DD (UTC, data do dia em que a Action rodou)
        geradoEm: now.toLocaleString('pt-BR'),
        kpis: {
            valor: Math.round(totalValor * 100) / 100,
            qtd: pedidosUnicos.size,
            clientes: clientesUnicos.size,
            diasMax: maxDiasGlobal,
        },
        regioes: regioesOrdenadas.map(label => {
            const r = porRegiao.get(label);
            return {
                label,
                valor: Math.round(r.valor * 100) / 100,
                qtd: r.pedidosSet.size,
                diasMax: r.diasMax,
            };
        }),
    };

    let historico = [];
    if (fs.existsSync(HISTORICO_ARQUIVO)) {
        try {
            historico = JSON.parse(fs.readFileSync(HISTORICO_ARQUIVO, 'utf-8'));
            if (!Array.isArray(historico)) historico = [];
        } catch (e) {
            historico = [];
        }
    }

    // Se já rodou hoje (ex.: reprocessamento), substitui a entrada de hoje em vez de duplicar.
    historico = historico.filter(h => h.data !== snapshot.data);
    historico.push(snapshot);
    historico.sort((a, b) => a.data.localeCompare(b.data));
    if (historico.length > HISTORICO_MAX_DIAS) {
        historico = historico.slice(historico.length - HISTORICO_MAX_DIAS);
    }

    fs.writeFileSync(HISTORICO_ARQUIVO, JSON.stringify(historico, null, 2), 'utf-8');
    console.log(`✅ Snapshot de ${snapshot.data} salvo. Histórico agora tem ${historico.length} dia(s).`);
    console.log(`   KPIs de hoje: valor=${snapshot.kpis.valor} qtd=${snapshot.kpis.qtd} clientes=${snapshot.kpis.clientes} diasMax=${snapshot.kpis.diasMax}`);
}

main();
