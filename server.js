const express = require("express");
const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");
const sharp = require("sharp");
const { DatabaseSync } = require("node:sqlite");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ANO_INICIAL = 2019;
const ANO_FINAL = 2025;
const caminhoBanco = path.join(__dirname, "data", "tce-pr.sqlite");

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));
app.get("/", (req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));

let bancoDados;
let municipiosCache = [];
let pronto = false;

function consultarBanco(sql, parametros = []) {
    if (!bancoDados) throw new Error("Banco de dados indisponível.");
    return bancoDados.prepare(sql).all(...parametros);
}

async function iniciarBancoDados() {
    if (!fs.existsSync(caminhoBanco)) throw new Error(`Banco não encontrado: ${caminhoBanco}`);
    bancoDados = new DatabaseSync(caminhoBanco, { readOnly: true });
    municipiosCache = consultarBanco(
        "SELECT municipio_id AS id, municipio_nome AS nome FROM municipios ORDER BY municipio_nome"
    ).map(item => ({ id: String(item.id), nome: String(item.nome) }));
    const total = Number(consultarBanco("SELECT COUNT(*) AS total FROM indicadores")[0]?.total || 0);
    if (municipiosCache.length !== 399) throw new Error(`Esperados 399 municípios; encontrados ${municipiosCache.length}.`);
    if (total !== 2793) throw new Error(`Esperados 2.793 registros; encontrados ${total}.`);
    pronto = true;
    console.log(`SQLite carregado: ${municipiosCache.length} municípios e ${total} registros.`);
}

function formatarBR(valor) {
    if (valor === null || valor === undefined || !Number.isFinite(Number(valor))) return "-";
    return Number(valor).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function montarIndicadoresBanco(registro) {
    const rclAjustada = Number(registro.rcl_ajustada);
    const dtp = Number(registro.dtp);
    const fundeb = Number(registro.fundeb);
    const percentual = rclAjustada > 0 ? dtp / rclAjustada * 100 : 0;
    const limiteAlertaPercentual = 48.6;
    const limitePrudencialPercentual = 51.3;
    const limiteMaximoPercentual = 54;
    let situacao = "OK";
    if (percentual >= limiteMaximoPercentual) situacao = "ACIMA DO LIMITE MÁXIMO";
    else if (percentual >= limitePrudencialPercentual) situacao = "LIMITE PRUDENCIAL";
    else if (percentual >= limiteAlertaPercentual) situacao = "LIMITE DE ALERTA";
    return {
        ano: Number(registro.ano), rclAjustada, rclAjustadaFormatada: formatarBR(rclAjustada),
        dtp, dtpFormatado: formatarBR(dtp), fundeb, fundebFormatado: formatarBR(fundeb),
        percentual, percentualFormatado: `${percentual.toFixed(2)}%`,
        limiteAlerta: rclAjustada * limiteAlertaPercentual / 100,
        limiteAlertaFormatado: formatarBR(rclAjustada * limiteAlertaPercentual / 100),
        limiteAlertaPercentual,
        limitePrudencial: rclAjustada * limitePrudencialPercentual / 100,
        limitePrudencialFormatado: formatarBR(rclAjustada * limitePrudencialPercentual / 100),
        limitePrudencialPercentual,
        limiteMaximo: rclAjustada * limiteMaximoPercentual / 100,
        limiteMaximoFormatado: formatarBR(rclAjustada * limiteMaximoPercentual / 100),
        limiteMaximoPercentual, situacao,
        fundebOrigem: "Banco histórico — RREO, Transferências do FUNDEB, últimos 12 meses (dezembro)"
    };
}

function obterHistoricoBanco(municipioId) {
    const registros = consultarBanco(
        `SELECT municipio_id, municipio_nome, ano, rcl_ajustada, dtp, fundeb
         FROM indicadores WHERE municipio_id = ? AND ano BETWEEN 2019 AND 2025 ORDER BY ano`,
        [String(municipioId)]
    );
    return registros.map(montarIndicadoresBanco).map((item, index, lista) => {
        const anterior = index ? lista[index - 1] : null;
        const base = lista[0];
        const evolucaoRCLAjustada = calcularEvolucao(item.rclAjustada, anterior?.rclAjustada);
        const acumuladoRCLAjustada = calcularEvolucaoAcumulada(item.rclAjustada, base.rclAjustada);
        const evolucaoDTP = calcularEvolucao(item.dtp, anterior?.dtp);
        const acumuladoDTP = calcularEvolucaoAcumulada(item.dtp, base.dtp);
        const evolucaoFundeb = calcularEvolucao(item.fundeb, anterior?.fundeb);
        const acumuladoFundeb = calcularEvolucaoAcumulada(item.fundeb, base.fundeb);
        return {
            ...item,
            evolucaoRCLAjustada, evolucaoRCLAjustadaFormatada: formatarPercentualEvolucao(evolucaoRCLAjustada),
            acumuladoRCLAjustada, acumuladoRCLAjustadaFormatado: formatarPercentualEvolucao(acumuladoRCLAjustada),
            evolucaoDTP, evolucaoDTPFormatada: formatarPercentualEvolucao(evolucaoDTP),
            acumuladoDTP, acumuladoDTPFormatado: formatarPercentualEvolucao(acumuladoDTP),
            evolucaoFundeb, evolucaoFundebFormatada: formatarPercentualEvolucao(evolucaoFundeb),
            acumuladoFundeb, acumuladoFundebFormatado: formatarPercentualEvolucao(acumuladoFundeb)
        };
    });
}

function calcularEvolucao(valorAtual, valorAnterior) {
    if (
        valorAtual === null ||
        valorAtual === undefined ||
        Number.isNaN(Number(valorAtual))
    ) {
        return null;
    }

    if (
        valorAnterior === null ||
        valorAnterior === undefined ||
        Number.isNaN(Number(valorAnterior)) ||
        Number(valorAnterior) === 0
    ) {
        return null;
    }

    return (
        (
            (Number(valorAtual) - Number(valorAnterior)) /
            Number(valorAnterior)
        ) * 100
    );
}

function calcularEvolucaoAcumulada(valorAtual, valorBase) {
    if (
        valorAtual === null ||
        valorAtual === undefined ||
        Number.isNaN(Number(valorAtual))
    ) {
        return null;
    }

    if (
        valorBase === null ||
        valorBase === undefined ||
        Number.isNaN(Number(valorBase)) ||
        Number(valorBase) === 0
    ) {
        return null;
    }

    return (
        (
            (Number(valorAtual) - Number(valorBase)) /
            Number(valorBase)
        ) * 100
    );
}

function formatarPercentualEvolucao(valor) {
    if (
        valor === null ||
        valor === undefined ||
        Number.isNaN(Number(valor))
    ) {
        return "-";
    }

    const numero = Number(valor);
    const sinal = numero > 0 ? "+" : "";

    return (
        sinal +
        numero.toFixed(2).replace(".", ",") +
        "%"
    );
}

// ======================================================
// HISTÓRICO
// ======================================================

function calcularHistorico(historico) {
    if (!historico || historico.length === 0) {
        return [];
    }

    const base =
        historico.find(item => Number(item.ano) === ANO_INICIAL) ||
        historico[0];

    return historico.map((item, index) => {
        const anterior =
            index > 0
                ? historico[index - 1]
                : null;

        const evolucaoRCLAjustada =
            calcularEvolucao(
                item.rclAjustada,
                anterior
                    ? anterior.rclAjustada
                    : null
            );

        const acumuladoRCLAjustada =
            calcularEvolucaoAcumulada(
                item.rclAjustada,
                base.rclAjustada
            );

        const evolucaoDTP =
            calcularEvolucao(
                item.dtp,
                anterior ? anterior.dtp : null
            );

        const acumuladoDTP =
            calcularEvolucaoAcumulada(
                item.dtp,
                base.dtp
            );

        const evolucaoFundeb =
            calcularEvolucao(
                item.fundeb,
                anterior ? anterior.fundeb : null
            );

        const acumuladoFundeb =
            calcularEvolucaoAcumulada(
                item.fundeb,
                base.fundeb
            );

        return {
            ...item,

            evolucaoRCLAjustada,
            evolucaoRCLAjustadaFormatada:
                formatarPercentualEvolucao(
                    evolucaoRCLAjustada
                ),

            acumuladoRCLAjustada,
            acumuladoRCLAjustadaFormatado:
                formatarPercentualEvolucao(
                    acumuladoRCLAjustada
                ),

            evolucaoDTP,
            evolucaoDTPFormatada:
                formatarPercentualEvolucao(
                    evolucaoDTP
                ),

            acumuladoDTP,
            acumuladoDTPFormatado:
                formatarPercentualEvolucao(
                    acumuladoDTP
                ),

            evolucaoFundeb,
            evolucaoFundebFormatada:
                formatarPercentualEvolucao(
                    evolucaoFundeb
                ),

            acumuladoFundeb,
            acumuladoFundebFormatado:
                formatarPercentualEvolucao(
                    acumuladoFundeb
                )
        };
    });
}

// ======================================================


app.get("/api/status", (req, res) => res.json({
    sucesso: true, pronto, municipios: municipiosCache.length, anoAtual: 2025, mesAtual: 12,
    anos: ["2019", "2020", "2021", "2022", "2023", "2024", "2025"]
}));
app.get("/api/municipios", (req, res) => res.json({ sucesso: true, municipios: municipiosCache }));

function obterRegistro(municipioId, ano) {
    return consultarBanco(
        `SELECT municipio_id, municipio_nome, ano, rcl_ajustada, dtp, fundeb
         FROM indicadores WHERE municipio_id = ? AND ano = ?`,
        [String(municipioId), Number(ano)]
    )[0];
}

app.post("/api/consulta", (req, res) => {
    try {
        const municipioId = String(req.body.municipio || "");
        const ano = Number(req.body.ano);
        if (!municipioId) throw new Error("Município não informado.");
        if (!Number.isInteger(ano) || ano < 2019 || ano > 2025) throw new Error("Ano inválido.");
        const registro = obterRegistro(municipioId, ano);
        if (!registro) throw new Error("Não existem dados para o município e ano selecionados.");
        res.json({ sucesso: true, dados: montarIndicadoresBanco(registro) });
    } catch (erro) { res.status(400).json({ sucesso: false, erro: erro.message }); }
});

app.get("/consultar", (req, res) => {
    try {
        const municipioId = String(req.query.municipio || "");
        const ano = Number(req.query.ano);
        if (!municipioId) throw new Error("Município não informado.");
        if (!Number.isInteger(ano) || ano < 2019 || ano > 2025) throw new Error("Ano inválido.");
        const registro = obterRegistro(municipioId, ano);
        if (!registro) throw new Error("Não existem dados para o município e ano selecionados.");
        res.json({ sucesso: true, dados: montarIndicadoresBanco(registro) });
    } catch (erro) { res.status(400).json({ sucesso: false, erro: erro.message }); }
});

let historicoExportCache = new Map();
app.get("/historico", (req, res) => {
    try {
        if (!pronto) throw new Error("O banco histórico ainda está sendo carregado.");
        const municipioId = String(req.query.municipio || "");
        if (!municipioId) throw new Error("Município não informado.");
        const municipio = municipiosCache.find(item => item.id === municipioId);
        if (!municipio) throw new Error("Município não encontrado no banco histórico.");
        const dados = obterHistoricoBanco(municipioId);
        if (dados.length !== 7) throw new Error(`O histórico contém ${dados.length} anos; eram esperados 7.`);
        historicoExportCache.set(municipioId, { dados, municipioNome: municipio.nome });
        res.json({ sucesso: true, anoInicial: 2019, anoFinal: 2025, periodo: "12", dados });
    } catch (erro) { res.status(400).json({ sucesso: false, erro: erro.message }); }
});

function corPercentualExcel(percentual) {
    const valor = Number(percentual);

    if (Number.isNaN(valor)) return null;

    if (valor >= 54) {
        return ["FFD71920", "FFFFFFFF"];
    }

    if (valor >= 51.3) {
        return ["FFF59E0B", "FFFFFFFF"];
    }

    if (valor >= 48.6) {
        return ["FFEAD800", "FF2C2C00"];
    }

    return ["FF55C83B", "FFFFFFFF"];
}

function formatarEvolucaoExcel(valor) {
    if (
        valor === null ||
        valor === undefined ||
        Number.isNaN(Number(valor))
    ) {
        return "-";
    }

    const numero = Number(valor);
    const sinal = numero > 0 ? "+" : "";

    return (
        sinal +
        numero.toFixed(2).replace(".", ",") +
        "%"
    );
}

function calcularEvolucaoExcel(atual, anterior) {
    if (
        atual === null ||
        atual === undefined ||
        anterior === null ||
        anterior === undefined ||
        Number.isNaN(Number(atual)) ||
        Number.isNaN(Number(anterior)) ||
        Number(anterior) === 0
    ) {
        return null;
    }

    return (
        ((Number(atual) - Number(anterior)) /
            Number(anterior)) *
        100
    );
}

function calcularAcumuladaExcel(atual, base) {
    if (
        atual === null ||
        atual === undefined ||
        base === null ||
        base === undefined ||
        Number.isNaN(Number(atual)) ||
        Number.isNaN(Number(base)) ||
        Number(base) === 0
    ) {
        return null;
    }

    return (
        ((Number(atual) - Number(base)) /
            Number(base)) *
        100
    );
}

function montarDadosExcel(historico) {
    const lista = Array.isArray(historico)
        ? historico
        : [];

    const anos = lista.map(
        item => String(item.ano)
    );

    const base =
        lista.find(
            item => Number(item.ano) === ANO_INICIAL
        ) || lista[0];

    const linhas = [
        ["Indicador", "Tipo", ...anos]
    ];

    const adicionarLinha = (
        indicador,
        tipo,
        valores
    ) => {
        linhas.push([
            indicador,
            tipo,
            ...valores
        ]);
    };

    const evolucaoAnual = (
        campo
    ) =>
        lista.map(
            (item, indice) =>
                formatarEvolucaoExcel(
                    indice === 0
                        ? null
                        : calcularEvolucaoExcel(
                              item[campo],
                              lista[indice - 1]?.[campo]
                          )
                )
        );

    const evolucaoAcumulada = (
        campo
    ) =>
        lista.map(
            (item, indice) =>
                formatarEvolucaoExcel(
                    indice === 0
                        ? null
                        : calcularAcumuladaExcel(
                              item[campo],
                              base?.[campo]
                          )
                )
        );

    adicionarLinha(
        "RGF — Receita Corrente Líquida Ajustada",
        "Valor",
        lista.map(
            item =>
                item.rclAjustadaFormatada ||
                formatarBR(item.rclAjustada)
        )
    );

    adicionarLinha(
        "Evolução anual",
        "%",
        evolucaoAnual("rclAjustada")
    );

    adicionarLinha(
        "Evolução acumulada",
        "%",
        evolucaoAcumulada("rclAjustada")
    );

    adicionarLinha(
        "RGF — Despesa Total com Pessoal",
        "Valor",
        lista.map(
            item =>
                item.dtpFormatado ||
                formatarBR(item.dtp)
        )
    );

    adicionarLinha(
        "Evolução anual",
        "%",
        evolucaoAnual("dtp")
    );

    adicionarLinha(
        "Evolução acumulada",
        "%",
        evolucaoAcumulada("dtp")
    );

    adicionarLinha(
        "RGF — Despesa Total com Pessoal %",
        "Valor",
        lista.map(
            item =>
                item.percentualFormatado ||
                `${Number(item.percentual || 0).toFixed(2)}%`
        )
    );

    adicionarLinha(
        "RREO — Transferências FUNDEB",
        "Valor",
        lista.map(
            item =>
                item.fundebFormatado ||
                formatarBR(item.fundeb)
        )
    );

    adicionarLinha(
        "Evolução anual",
        "%",
        evolucaoAnual("fundeb")
    );

    adicionarLinha(
        "Evolução acumulada",
        "%",
        evolucaoAcumulada("fundeb")
    );

    return {
        lista,
        anos,
        linhas
    };
}

async function criarImagemGraficoExcel(lista) {
    const largura = 1000;
    const altura = 520;
    const esquerda = 105;
    const direita = 35;
    const topo = 92;
    const base = 62;
    const larguraGrafico =
        largura - esquerda - direita;
    const alturaGrafico =
        altura - topo - base;

    const maximo = Math.max(
        1,
        ...lista.flatMap(item => [
            Number(item.rclAjustada) || 0,
            Number(item.fundeb) || 0,
            Number(item.dtp) || 0
        ])
    );

    const teto = maximo * 1.12;
    const quantidade =
        Math.max(lista.length, 1);
    const grupo =
        larguraGrafico / quantidade;
    const barra = Math.min(
        30,
        grupo * 0.18
    );
    const espacamento = barra * 0.22;

    const y = valor =>
        topo +
        alturaGrafico -
        ((Number(valor) || 0) / teto) *
            alturaGrafico;

    const grade = [];
    const eixo = [];

    for (let i = 0; i <= 5; i++) {
        const valor =
            (teto / 5) * i;
        const yy = y(valor);

        grade.push(`
            <line
                x1="${esquerda}"
                y1="${yy}"
                x2="${largura - direita}"
                y2="${yy}"
                stroke="#dfe7ef"
                stroke-width="1"
            />
        `);

        eixo.push(`
            <text
                x="${esquerda - 14}"
                y="${yy + 5}"
                text-anchor="end"
                font-family="Arial"
                font-size="18"
                fill="#617286"
            >${formatarBR(Math.round(valor))}</text>
        `);
    }

    const barras = [];
    const anos = [];

    lista.forEach(
        (item, indice) => {
            const centro =
                esquerda +
                grupo * indice +
                grupo / 2;

            [
                [
                    item.rclAjustada,
                    "#2477a8",
                    -barra - espacamento
                ],
                [
                    item.fundeb,
                    "#30964f",
                    0
                ],
                [
                    item.dtp,
                    "#d44949",
                    barra + espacamento
                ]
            ].forEach(
                ([valor, cor, deslocamento]) => {
                    const topoBarra = y(valor);
                    const alturaBarra =
                        topo +
                        alturaGrafico -
                        topoBarra;

                    barras.push(`
                        <rect
                            x="${centro + deslocamento - barra / 2}"
                            y="${topoBarra}"
                            width="${barra}"
                            height="${Math.max(0, alturaBarra)}"
                            fill="${cor}"
                        />
                    `);
                }
            );

            anos.push(`
                <text
                    x="${centro}"
                    y="${altura - 48}"
                    text-anchor="middle"
                    font-family="Arial"
                    font-size="19"
                    font-weight="700"
                    fill="#617286"
                >${String(item.ano)}</text>
            `);
        }
    );

    const svg = `
        <svg
            xmlns="http://www.w3.org/2000/svg"
            width="${largura}"
            height="${altura}"
        >
            <rect
                width="100%"
                height="100%"
                fill="#ffffff"
            />

            <text
                x="${esquerda}"
                y="34"
                font-family="Arial"
                font-size="22"
                font-weight="700"
                fill="#122b42"
            >Evolução dos principais indicadores financeiros</text>

            <rect x="105" y="46" width="24" height="12" fill="#2477a8"/>
            <text x="139" y="58" font-family="Arial" font-size="16" fill="#33465d">RCL Ajustada</text>

            <rect x="300" y="46" width="24" height="12" fill="#30964f"/>
            <text x="334" y="58" font-family="Arial" font-size="16" fill="#33465d">FUNDEB</text>

            <rect x="470" y="46" width="24" height="12" fill="#d44949"/>
            <text x="504" y="58" font-family="Arial" font-size="16" fill="#33465d">DTP</text>

            ${grade.join("")}
            ${eixo.join("")}

            <line
                x1="${esquerda}"
                y1="${topo + alturaGrafico}"
                x2="${largura - direita}"
                y2="${topo + alturaGrafico}"
                stroke="#b9c6d3"
                stroke-width="1"
            />

            ${barras.join("")}
            ${anos.join("")}
        </svg>
    `;

    return sharp(
        Buffer.from(svg)
    )
        .png()
        .toBuffer();
}

async function gerarExcelFormatado(
    historico,
    municipioNome
) {
    const {
        lista,
        anos,
        linhas
    } = montarDadosExcel(historico);

    const workbook =
        new ExcelJS.Workbook();

    workbook.creator =
        "Automação TCE-PR";
    workbook.title =
        "Análise TCE-PR";
    workbook.subject =
        "Histórico financeiro municipal";
    workbook.created =
        new Date();

    const worksheet =
        workbook.addWorksheet(
            "Histórico 2019-2025",
            {
                views: [
                    {
                        state: "frozen",
                        ySplit: 5,
                        xSplit: 2,
                        zoomScale: 90
                    }
                ]
            }
        );

    const ultimaColuna =
        String.fromCharCode(
            64 + anos.length + 2
        );

    worksheet.mergeCells(
        `A1:${ultimaColuna}1`
    );
    worksheet.mergeCells(
        `A2:${ultimaColuna}2`
    );
    worksheet.mergeCells(
        `A3:${ultimaColuna}3`
    );

    worksheet.getCell("A1").value =
        "ANÁLISE TCE-PR — HISTÓRICO 2019–2025";

    worksheet.getCell("A1").font = {
        name: "Arial",
        size: 18,
        bold: true,
        color: {
            argb: "FFFFFFFF"
        }
    };

    worksheet.getCell("A1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
            argb: "FF123B5D"
        }
    };

    worksheet.getCell("A1").alignment = {
        vertical: "middle",
        horizontal: "left"
    };

    worksheet.getRow(1).height = 34;

    worksheet.getCell("A2").value =
        `Município: ${municipioNome} • Período: dezembro • Anos: ${anos.join(", ")}`;

    worksheet.getCell("A2").font = {
        name: "Arial",
        size: 11,
        bold: true,
        color: {
            argb: "FF176DA8"
        }
    };

    worksheet.getCell("A2").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
            argb: "FFEAF5FF"
        }
    };

    worksheet.getCell("A2").alignment = {
        vertical: "middle",
        horizontal: "left"
    };

    worksheet.getRow(2).height = 24;

    worksheet.getCell("A3").value =
        "Histórico financeiro municipal com evolução anual, evolução acumulada, percentual da DTP e referência visual dos limites da LRF.";

    worksheet.getCell("A3").font = {
        name: "Arial",
        size: 10,
        italic: true,
        color: {
            argb: "FF66788C"
        }
    };

    worksheet.getCell("A3").alignment = {
        vertical: "middle",
        horizontal: "left",
        wrapText: true
    };

    worksheet.getRow(3).height = 30;

    linhas.forEach(
        (linha, indice) => {
            const row =
                worksheet.addRow(linha);

            row.height =
                indice === 0
                    ? 32
                    : 27;

            row.eachCell(
                (cell, coluna) => {
                    cell.border = {
                        top: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        },
                        left: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        },
                        bottom: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        },
                        right: {
                            style: "thin",
                            color: {
                                argb: "FFDCE5ED"
                            }
                        }
                    };

                    cell.alignment = {
                        vertical: "middle",
                        horizontal:
                            coluna <= 2
                                ? "left"
                                : "right",
                        wrapText: true
                    };

                    if (indice === 0) {
                        cell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: {
                                argb: "FF176DA8"
                            }
                        };

                        cell.font = {
                            name: "Arial",
                            size: 11,
                            bold: true,
                            color: {
                                argb: "FFFFFFFF"
                            }
                        };

                        cell.alignment = {
                            vertical: "middle",
                            horizontal:
                                coluna <= 2
                                    ? "left"
                                    : "center"
                        };
                    } else {
                        cell.font = {
                            name: "Arial",
                            size: 10,
                            bold: coluna <= 2,
                            color: {
                                argb: "FF162B40"
                            }
                        };

                        if (coluna <= 2) {
                            cell.fill = {
                                type: "pattern",
                                pattern: "solid",
                                fgColor: {
                                    argb:
                                        indice % 2 === 0
                                            ? "FFF8FAFC"
                                            : "FFFFFFFF"
                                }
                            };
                        }
                    }
                }
            );

            if (
                indice > 0 &&
                linha[1] === "%"
            ) {
                for (
                    let coluna = 3;
                    coluna <= linha.length;
                    coluna++
                ) {
                    const texto =
                        row.getCell(coluna).value;

                    if (
                        typeof texto !== "string" ||
                        texto === "-"
                    ) {
                        continue;
                    }

                    const valor =
                        Number(
                            texto
                                .replace("+", "")
                                .replace("%", "")
                                .replace(",", ".")
                        );

                    if (Number.isNaN(valor)) {
                        continue;
                    }

                    const positivo =
                        valor > 0;
                    const negativo =
                        valor < 0;

                    row.getCell(coluna).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: {
                            argb: positivo
                                ? "FFE8F5E9"
                                : negativo
                                    ? "FFFFEEEE"
                                    : "FFF3F6FA"
                        }
                    };

                    row.getCell(coluna).font = {
                        name: "Arial",
                        size: 10,
                        bold: true,
                        color: {
                            argb: positivo
                                ? "FF217A19"
                                : negativo
                                    ? "FFD71920"
                                    : "FF66788C"
                        }
                    };
                }
            }

            if (
                indice > 0 &&
                linha[0] ===
                    "RGF — Despesa Total com Pessoal %"
            ) {
                for (
                    let coluna = 3;
                    coluna <= linha.length;
                    coluna++
                ) {
                    const texto =
                        row.getCell(coluna).value;

                    const valor =
                        Number(
                            String(texto)
                                .replace("%", "")
                                .replace(",", ".")
                        );

                    const cor =
                        corPercentualExcel(valor);

                    if (!cor) continue;

                    row.getCell(coluna).fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: {
                            argb: cor[0]
                        }
                    };

                    row.getCell(coluna).font = {
                        name: "Arial",
                        size: 10,
                        bold: true,
                        color: {
                            argb: cor[1]
                        }
                    };

                    row.getCell(coluna).alignment = {
                        vertical: "middle",
                        horizontal: "center"
                    };
                }
            }
        }
    );

    worksheet.getColumn(1).width = 52;
    worksheet.getColumn(2).width = 14;

    for (
        let coluna = 3;
        coluna <= Math.max(anos.length + 2, 12);
        coluna++
    ) {
        worksheet.getColumn(coluna).width = 23;
    }

    worksheet.pageSetup = {
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        paperSize:
            worksheet.PAPERSIZE_A4
    };

    worksheet.getColumn(10).width = 16;
    worksheet.getColumn(11).width = 27;
    worksheet.getColumn(12).width = 3;

    worksheet.mergeCells("J5:K5");

    const tituloLegenda =
        worksheet.getCell("J5");

    tituloLegenda.value =
        "Legenda — Percentual DTP";

    tituloLegenda.font = {
        name: "Arial",
        size: 11,
        bold: true,
        color: {
            argb: "FFFFFFFF"
        }
    };

    tituloLegenda.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
            argb: "FF176DA8"
        }
    };

    tituloLegenda.alignment = {
        vertical: "middle",
        horizontal: "left"
    };

    worksheet.getRow(5).height = 30;

    [
        [
            "Até 48,6%",
            "Situação regular",
            "FF55C83B"
        ],
        [
            "48,6% a 51,3%",
            "Alerta",
            "FFEAD800"
        ],
        [
            "51,3% a 54%",
            "Limite prudencial",
            "FFF59E0B"
        ],
        [
            "Acima de 54%",
            "Acima do limite legal",
            "FFD71920"
        ]
    ].forEach(
        (item, indice) => {
            const rowNumero = 6 + indice;
            const row = worksheet.getRow(rowNumero);

            row.height = 24;

            const faixa = row.getCell(10);
            const descricao = row.getCell(11);

            faixa.value = item[0];
            descricao.value = item[1];

            faixa.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: {
                    argb: item[2]
                }
            };

            faixa.font = {
                name: "Arial",
                size: 9,
                bold: true,
                color: {
                    argb:
                        item[2] === "FFEAD800"
                            ? "FF2C2C00"
                            : "FFFFFFFF"
                }
            };

            descricao.font = {
                name: "Arial",
                size: 9,
                bold: true,
                color: {
                    argb: "FF33465D"
                }
            };

            faixa.alignment = {
                vertical: "middle",
                horizontal: "center",
                wrapText: true
            };

            descricao.alignment = {
                vertical: "middle",
                horizontal: "left",
                wrapText: true
            };

            [faixa, descricao].forEach(cell => {
                cell.border = {
                    top: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    },
                    left: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    },
                    bottom: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    },
                    right: {
                        style: "thin",
                        color: {
                            argb: "FFDCE5ED"
                        }
                    }
                };
            });
        }
    );

    const grafico =
        workbook.addWorksheet(
            "Gráfico"
        );

    grafico.views = [
        {
            showGridLines: false,
            zoomScale: 85
        }
    ];

    grafico.getCell("A1").value =
        `Evolução dos principais indicadores — ${municipioNome}`;

    grafico.getCell("A1").font = {
        name: "Arial",
        size: 18,
        bold: true,
        color: {
            argb: "FF122B42"
        }
    };

    const imagem =
        await criarImagemGraficoExcel(
            lista
        );

    grafico.getRow(1).height = 28;

    for (let coluna = 1; coluna <= 12; coluna++) {
        grafico.getColumn(coluna).width = 10;
    }

    const imagemId =
        workbook.addImage({
            buffer: imagem,
            extension: "png"
        });

    grafico.addImage(
        imagemId,
        {
            tl: {
                col: 0,
                row: 2
            },
            ext: {
                width: 900,
                height: 468
            }
        }
    );

    return workbook.xlsx.writeBuffer();
}


app.get("/api/exportar-excel", async (req, res) => {
    try {
        const municipioId = String(req.query.municipio || "");
        if (!municipioId) return res.status(400).json({ sucesso: false, erro: "Município não informado." });
        const cache = historicoExportCache.get(municipioId);
        if (!cache?.dados?.length) return res.status(400).json({ sucesso: false, erro: "Consulte o histórico antes de exportar para o Excel." });
        const arquivo = await gerarExcelFormatado(cache.dados, cache.municipioNome);
        const nome = cache.municipioNome.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "_");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename="Historico_${nome}_2019-2025.xlsx"`);
        res.end(arquivo);
    } catch (erro) {
        console.error("ERRO AO GERAR EXCEL:", erro);
        res.status(500).json({ sucesso: false, erro: "Não foi possível gerar o arquivo Excel formatado." });
    }
});

(async () => {
    try {
        await iniciarBancoDados();
        app.listen(PORT, () => console.log(`\nAutomação TCE-PR — consulta SQLite\nServidor: http://localhost:${PORT}\nPeríodo: 2019–2025\nMunicípios: ${municipiosCache.length}\n`));
    } catch (erro) {
        console.error("ERRO AO INICIAR:", erro.message);
        process.exit(1);
    }
})();

process.on("SIGINT", () => {
    if (bancoDados) bancoDados.close();
    process.exit(0);
});
