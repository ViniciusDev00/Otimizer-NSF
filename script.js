pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

let arquivosPdf = [];
let dadosMondayExcel = null;
let moduloAtivo = 'separacao'; // 'separacao' ou 'carregamento'

// NAVEGAÇÃO ENTRE MÓDULOS NA TELA INICIAL
window.selecionarModulo = function(modulo) {
    moduloAtivo = modulo;
    document.getElementById('tela-modulo').style.display = 'none';
    document.getElementById('painel-operacao').style.display = 'block';

    const badge = document.getElementById('badge-modulo-ativo');
    const secSeparacao = document.getElementById('opcoes-separacao');
    const secCarregamento = document.getElementById('opcoes-carregamento');

    if (modulo === 'separacao') {
        badge.innerText = '📦 MÓDULO SEPARAÇÃO';
        badge.style.backgroundColor = '#2563eb';
        secSeparacao.style.display = 'block';
        secCarregamento.style.display = 'none';
    } else {
        badge.innerText = '🚚 MÓDULO CARREGAMENTO / EXPEDIÇÃO';
        badge.style.backgroundColor = '#d97706';
        secSeparacao.style.display = 'none';
        secCarregamento.style.display = 'block';
    }
};

window.voltarParaModulos = function() {
    document.getElementById('painel-operacao').style.display = 'none';
    document.getElementById('tela-modulo').style.display = 'block';
};

// HANDLERS DE UPLOAD DE ARQUIVOS
document.getElementById('input-pdf').addEventListener('change', function(e) {
    arquivosPdf = Array.from(e.target.files);
    document.getElementById('lista-pdfs').innerHTML = arquivosPdf.length > 0 
        ? `✔️ ${arquivosPdf.length} PDF(s) carregado(s).` 
        : "Nenhum PDF selecionado";
    verificarRequisitos();
});

document.getElementById('input-monday').addEventListener('change', function(e) {
    const arquivo = e.target.files[0];
    if (arquivo) {
        const leitor = new FileReader();
        leitor.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const primeiraAba = workbook.SheetNames[0];
            const planilha = workbook.Sheets[primeiraAba];
            dadosMondayExcel = XLSX.utils.sheet_to_json(planilha, { header: 1 });
            document.getElementById('status-monday').innerHTML = `✔️ Planilha "${arquivo.name}" carregada.`;
            verificarRequisitos();
        };
        leitor.readAsArrayBuffer(arquivo);
    }
});

function verificarRequisitos() {
    const btnProcessar = document.getElementById('btn-processar');
    btnProcessar.disabled = !(arquivosPdf.length > 0 && dadosMondayExcel !== null);
}

// BUSCA NO EXCEL COM MAPEAMENTO RIGOROSO DE CABEÇALHO
function consultarDadosMonday(numeroPedido) {
    let resultado = {
        aviso: "Não Encontrado",
        osNum: "Não Encontrado",
        tipo: "Padrão",
        cidadeEstado: "Não Informado",
        transporte: "Caminhão",
        mercado: ""
    };

    if (!dadosMondayExcel || !numeroPedido) return resultado;
    let pedidoProcurado = numeroPedido.trim().toUpperCase();

    // 1. LOCALIZA A LINHA DE CABEÇALHO OFICIAL
    let linhaCabecalho = -1;
    for (let r = 0; r < Math.min(5, dadosMondayExcel.length); r++) {
        let linha = dadosMondayExcel[r];
        if (!linha) continue;
        let strLinha = linha.map(x => String(x || '').toUpperCase()).join(' ');
        if (strLinha.includes("Nº OS") || strLinha.includes("AVISO") || strLinha.includes("NOME DO MERCADO")) {
            linhaCabecalho = r;
            break;
        }
    }

    let idxAviso = 4;   // Coluna E
    let idxOS = 3;      // Coluna D (Nº OS)
    let idxTipo = 6;    // Coluna G
    let idxCidade = 12; // Coluna M
    let idxTransp = 10; // Coluna K (Possivel Transp)
    let idxMercado = 8; // Coluna I (Nome do Mercado)

    // 2. EXTRAI OS ÍNDICES APENAS DA LINHA DE CABEÇALHO
    if (linhaCabecalho !== -1) {
        let headerRow = dadosMondayExcel[linhaCabecalho];
        for (let c = 0; c < headerRow.length; c++) {
            let val = String(headerRow[c] || '').trim().toUpperCase();
            if (val === "AVISO" || val === "Nº AVISO") idxAviso = c;
            else if (val === "Nº OS" || val === "NUMERO OS" || val === "NÚMERO OS") idxOS = c;
            else if (val === "TIPO") idxTipo = c;
            else if (val.includes("CIDADE") || val.includes("ESTADO")) idxCidade = c;
            else if (val.includes("TRANSP") || val.includes("SEDEX")) idxTransp = c;
            else if (val === "NOME DO MERCADO" || val === "NOME MERCADO" || val === "CLIENTE" || val === "NOME DO CLIENTE") idxMercado = c;
        }
    }

    // 3. VARRE APENAS AS LINHAS DE DADOS ABAIXO DO CABEÇALHO
    let inicioDados = linhaCabecalho !== -1 ? linhaCabecalho + 1 : 1;

    for (let i = inicioDados; i < dadosMondayExcel.length; i++) {
        let linha = dadosMondayExcel[i];
        if (!linha || linha.length === 0) continue;

        let achou = false;
        for (let j = 0; j < linha.length; j++) {
            let valorCelula = String(linha[j] || '').trim().toUpperCase();
            if (valorCelula === pedidoProcurado || (pedidoProcurado.length > 3 && valorCelula.includes(pedidoProcurado))) {
                achou = true;
                break;
            }
        }

        if (achou) {
            if (idxAviso < linha.length && linha[idxAviso]) {
                let valAviso = String(linha[idxAviso]).trim();
                if (valAviso && !['nan', 'null', 'undefined'].includes(valAviso.toLowerCase())) {
                    let m = valAviso.match(/\d+/);
                    resultado.aviso = m ? m[0] : valAviso;
                }
            }

            if (idxOS < linha.length && linha[idxOS]) {
                let valOS = String(linha[idxOS]).trim();
                if (valOS && !['nan', 'null', 'undefined'].includes(valOS.toLowerCase())) {
                    resultado.osNum = valOS;
                }
            }

            if (idxTipo < linha.length && linha[idxTipo]) {
                let valTipo = String(linha[idxTipo]).trim();
                if (valTipo && !['nan', 'null', 'undefined'].includes(valTipo.toLowerCase())) {
                    resultado.tipo = valTipo;
                }
            }

            if (idxCidade < linha.length && linha[idxCidade]) {
                let valCidade = String(linha[idxCidade]).trim();
                if (valCidade && !['nan', 'null', 'undefined'].includes(valCidade.toLowerCase())) {
                    resultado.cidadeEstado = valCidade;
                }
            }

            if (idxTransp < linha.length && linha[idxTransp]) {
                let valTransp = String(linha[idxTransp]).trim();
                if (valTransp && !['nan', 'null', 'undefined'].includes(valTransp.toLowerCase())) {
                    resultado.transporte = valTransp;
                }
            }

            if (idxMercado < linha.length && linha[idxMercado]) {
                let valMercado = String(linha[idxMercado]).trim();
                if (valMercado && !['nan', 'null', 'undefined'].includes(valMercado.toLowerCase())) {
                    resultado.mercado = valMercado;
                }
            }

            break;
        }
    }
    return resultado;
}

// HIGIENIZADOR INTELIGENTE DE CLIENTE
function limparNomeCliente(nomeBruto) {
    if (!nomeBruto) return "";
    let nome = String(nomeBruto).trim();

    // Rejeita links HTTP e anexos de imagem/PDF
    if (nome.toLowerCase().startsWith('http') || nome.toLowerCase().includes('monday.com') || /\.(pdf|png|jpg|jpeg|webp)$/i.test(nome)) {
        return "";
    }

    nome = nome.toUpperCase();
    nome = nome.replace(/^CLIENTE\s*:\s*/i, '');
    nome = nome.replace(/^[0-9]{3,8}\s*(-\s*)?/, '');
    nome = nome.replace(/\(\d+\)/g, '');
    nome = nome.replace(/[\.\s,/-]*\b(LTDA|S\.A\.|SA|S\/A|ME|EPP|FILIAL\d*)\b.*/g, '');
    nome = nome.replace(/["']/g, '').trim();
    nome = nome.replace(/\s+/g, ' ');

    return nome;
}

// PROCESSADOR PRINCIPAL
document.getElementById('btn-processar').addEventListener('click', async function() {
    const btnImprimir = document.getElementById('btn-imprimir');
    const containerResultado = document.getElementById('resultado-impressao');
    const mapaAgrupadoGeral = {};

    for (let arquivo of arquivosPdf) {
        try {
            const arrayBuffer = await arquivo.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            
            let nrPedido = "";
            let matchPedidoId = arquivo.name.match(/P\d+/i);
            if (matchPedidoId) nrPedido = matchPedidoId[0].toUpperCase();
            
            let nomeClienteAlternativo = arquivo.name
                .replace(/P\d+/i, '').replace(/\.pdf$/i, '').replace(/[-_()]/g, ' ').trim();

            for (let i = 1; i <= pdf.numPages; i++) {
                const pagina = await pdf.getPage(i);
                const conteudoTexto = await pagina.getTextContent();
                const linhasTexto = conteudoTexto.items.map(item => item.str.trim()).filter(s => s !== "");

                if (!nrPedido) {
                    let idxPed = linhasTexto.findIndex(t => t.toUpperCase().includes("PEDIDO:"));
                    if (idxPed !== -1 && linhasTexto[idxPed + 1]) {
                        nrPedido = linhasTexto[idxPed + 1].replace(/["']/g, '').trim().toUpperCase();
                    }
                }

                let nomeClienteBruto = "";
                let linhaComCodigo = linhasTexto.find(t => /^\d+\s*-\s*/.test(t));
                if (linhaComCodigo) {
                    let nomeLimpo = linhaComCodigo.replace(/["']/g, '').trim();
                    if (!nomeLimpo.includes("- SP") && !nomeLimpo.includes("- MG") && nomeLimpo.length > 2) {
                        nomeClienteBruto = nomeLimpo;
                    }
                }

                if (!nomeClienteBruto || nomeClienteBruto === "DESCONHECIDO") {
                    nomeClienteBruto = nomeClienteAlternativo || "CLIENTE DESCONHECIDO";
                }
                
                let dadosMonday = consultarDadosMonday(nrPedido);
                
                let nomeDaPlanilha = limparNomeCliente(dadosMonday.mercado);
                let nomeDoPdf = limparNomeCliente(nomeClienteBruto);

                let nomeClienteFinal = nomeDaPlanilha ? nomeDaPlanilha : (nomeDoPdf ? nomeDoPdf : "CLIENTE DESCONHECIDO");

                for (let j = 0; j < linhasTexto.length; j++) {
                    if (/^\d+,\d{2}$/.test(linhasTexto[j])) {
                        let qtd = parseFloat(linhasTexto[j].replace(',', '.'));
                        
                        let partesDescricao = [];
                        let k = j - 1;
                        
                        while (k >= 0) {
                            let txt = linhasTexto[k].replace(/["']/g, '').trim();
                            let txtUpper = txt.toUpperCase();
                            
                            if (txtUpper === "PRODUTO" || txtUpper === "DESCRIÇÃO" || txtUpper === "LOCALIZAÇÃO" || txtUpper === "QTD" || /^\d+,\d{2}$/.test(txt)) {
                                break;
                            }
                            partesDescricao.unshift(txt);
                            if (/\b[A-Z0-9]{5,}\b/.test(txtUpper) && partesDescricao.length >= 2) break;
                            k--;
                        }

                        let descricaoCompleta = partesDescricao.join(' ').trim().replace(/^[0-9]{4,8}\s+/, '');
                        if (!descricaoCompleta) continue;

                        let chaveAgrupamento = `${nomeClienteFinal}_${dadosMonday.aviso}_${descricaoCompleta}`;

                        if (mapaAgrupadoGeral[chaveAgrupamento]) {
                            mapaAgrupadoGeral[chaveAgrupamento].qtd += qtd;
                            if (nrPedido && !mapaAgrupadoGeral[chaveAgrupamento].ordens.includes(nrPedido)) {
                                mapaAgrupadoGeral[chaveAgrupamento].ordens.push(nrPedido);
                            }
                        } else {
                            mapaAgrupadoGeral[chaveAgrupamento] = {
                                cliente: nomeClienteFinal,
                                aviso: dadosMonday.aviso,
                                osNum: dadosMonday.osNum,
                                tipo: dadosMonday.tipo,
                                cidadeEstado: dadosMonday.cidadeEstado,
                                transporte: dadosMonday.transporte,
                                ordens: nrPedido ? [nrPedido] : ["-"],
                                descricao: descricaoCompleta,
                                qtd: qtd
                            };
                        }
                    }
                }
            }
        } catch (erro) {
            console.error("Erro no arquivo: " + arquivo.name, erro);
        }
    }

    if (Object.keys(mapaAgrupadoGeral).length === 0) {
        alert("Nenhum dado extraído.");
        return;
    }

    const registrosOrdenados = Object.values(mapaAgrupadoGeral);

    let htmlFinal = "";

    // ==========================================
    // RENDERIZAÇÃO 1: MÓDULO DE CARREGAMENTO
    // ==========================================
    if (moduloAtivo === 'carregamento') {
        let carregamentoUnico = {};
        registrosOrdenados.forEach(reg => {
            let osFormatada = (reg.osNum && reg.osNum !== "Não Encontrado") ? reg.osNum : reg.ordens.join(', ');
            let chave = `${reg.cliente}_${reg.aviso}_${osFormatada}`;
            if (!carregamentoUnico[chave]) {
                carregamentoUnico[chave] = {
                    cliente: reg.cliente,
                    cidadeEstado: reg.cidadeEstado,
                    osNum: osFormatada,
                    aviso: reg.aviso,
                    tipo: reg.tipo,
                    transporte: reg.transporte
                };
            }
        });

        let listaCarregamento = Object.values(carregamentoUnico);
        listaCarregamento.sort((a, b) => a.cliente.localeCompare(b.cliente) || a.aviso.localeCompare(b.aviso));

        let contSedex = listaCarregamento.filter(item => item.transporte.toLowerCase().includes('sedex')).length;
        let contPicape = listaCarregamento.filter(item => item.transporte.toLowerCase().includes('picape') || item.transporte.toLowerCase().includes('carro')).length;
        let contCaminhao = listaCarregamento.filter(item => item.transporte.toLowerCase().includes('caminhão') || item.transporte.toLowerCase().includes('caminhao')).length;

        htmlFinal = `
            <div class="header-carregamento">
                <h3>🚚 PROGRAMAÇÃO DE CARREGAMENTO & EXPEDIÇÃO</h3>
                <div style="font-size: 10px; font-weight: 700;">
                    Emissão: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'})}
                </div>
            </div>

            <div class="box-modal-transporte">
                <div><strong>RESUMO DE EXPEDIÇÃO:</strong> <strong>${listaCarregamento.length}</strong> Carga(s) / Aviso(s)</div>
                <div><strong>MODAIS:</strong> ⚡ Sedex: <strong>${contSedex}</strong> | 🛻 Picape/Carro: <strong>${contPicape}</strong> | 🚚 Caminhão: <strong>${contCaminhao}</strong></div>
            </div>

            <table class="tabela-carregamento">
                <thead>
                    <tr>
                        <th style="width: 25%;">Cidade / Estado</th>
                        <th style="width: 20%;">N° OS</th>
                        <th style="width: 15%;">N° Aviso</th>
                        <th style="width: 20%;">Tipo Carga</th>
                        <th style="width: 20%;">Modo Transporte</th>
                    </tr>
                </thead>
                <tbody>
        `;

        let clienteAtual = "";

        listaCarregamento.forEach(reg => {
            if (reg.cliente !== clienteAtual) {
                clienteAtual = reg.cliente;
                htmlFinal += `
                    <tr class="secao-cliente">
                        <td colspan="5">👤 CLIENTE: <strong>${clienteAtual}</strong></td>
                    </tr>
                `;
            }

            let tUpper = reg.transporte.toUpperCase();
            let badgeTransp = `<span class="badge-transp-caminhao">🚚 CAMINHÃO</span>`;

            if (tUpper.includes('SEDEX')) {
                badgeTransp = `<span class="badge-transp-sedex">⚡ SEDEX</span>`;
            } else if (tUpper.includes('PICAPE') || tUpper.includes('CARRO')) {
                badgeTransp = `<span class="badge-transp-picape">🛻 CARRO / PICAPE</span>`;
            }

            htmlFinal += `
                <tr>
                    <td>${reg.cidadeEstado}</td>
                    <td><strong>${reg.osNum}</strong></td>
                    <td><strong>${reg.aviso}</strong></td>
                    <td><span class="badge-tipo">${reg.tipo}</span></td>
                    <td>${badgeTransp}</td>
                </tr>
            `;
        });

        htmlFinal += `</tbody></table>`;

    // ==========================================
    // RENDERIZAÇÃO 2: MÓDULO DE SEPARAÇÃO
    // ==========================================
    } else {
        registrosOrdenados.sort((a, b) => a.cliente.localeCompare(b.cliente) || a.descricao.localeCompare(b.descricao));
        const modoRelatorio = document.querySelector('input[name="opcao-modo"]:checked').value;
        if (modoRelatorio === "corporativo") document.body.classList.add("modo-corporativo");

        let totalPecasGeral = registrosOrdenados.reduce((acc, item) => acc + item.qtd, 0);

        htmlFinal = `
            <div class="header-hyper">
                <h4>⚡ LISTA GERENCIAL DE SEPARAÇÃO CONSOLIDADA</h4>
                <div class="meta">
                    Emissão: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'})} | 
                    <strong>Peças: <span id="lbl-total-pecas">${totalPecasGeral}</span></strong> | 
                    <strong>Itens: <span id="lbl-total-itens">${registrosOrdenados.length}</span></strong>
                </div>
            </div>

            <table class="tabela-hyper">
                <thead>
                    <tr>
                        <th class="col-descarte no-print" title="Marque para descartar da impressão">❌</th>
                        <th style="width: 12%;">N° Aviso</th>
                        <th style="width: 14%;">Ordem(ns)</th>
                        <th style="width: 68%;">Descrição do Item / Peça</th>
                        <th style="width: 6%; text-align: center;">Qtd</th>
                    </tr>
                </thead>
                <tbody>
        `;

        let clienteAtual = "";

        registrosOrdenados.forEach((reg, index) => {
            if (reg.cliente !== clienteAtual) {
                clienteAtual = reg.cliente;
                htmlFinal += `
                    <tr class="secao-cliente">
                        <td colspan="5">👤 CLIENTE: <strong>${clienteAtual}</strong></td>
                    </tr>
                `;
            }

            htmlFinal += `
                <tr id="row-item-${index}">
                    <td class="col-descarte no-print">
                        <input type="checkbox" class="chk-descarte" title="Descartar item volumado" onchange="alternarDescarteItem(${index})">
                    </td>
                    <td><strong>${reg.aviso}</strong></td>
                    <td>${reg.ordens.join(', ')}</td>
                    <td>${reg.descricao}</td>
                    <td style="text-align: center; font-weight: bold;" class="col-qtd">${reg.qtd}</td>
                </tr>
            `;
        });

        htmlFinal += `</tbody></table>`;
    }

    containerResultado.innerHTML = htmlFinal;
    btnImprimir.disabled = false;
});

window.alternarDescarteItem = function(index) {
    const row = document.getElementById(`row-item-${index}`);
    if (row) {
        row.classList.toggle('item-descartado');
        atualizarTotaisRelatorio();
    }
};

function atualizarTotaisRelatorio() {
    const linhasVisiveis = document.querySelectorAll('.tabela-hyper tbody tr:not(.secao-cliente):not(.item-descartado)');
    let totalPecas = 0;
    linhasVisiveis.forEach(row => {
        let celulaQtd = row.querySelector('.col-qtd');
        if (celulaQtd) totalPecas += parseFloat(celulaQtd.innerText || 0);
    });
    const elPecas = document.getElementById('lbl-total-pecas');
    const elItens = document.getElementById('lbl-total-itens');
    if (elPecas) elPecas.innerText = totalPecas;
    if (elItens) elItens.innerText = linhasVisiveis.length;
}

document.getElementById('btn-imprimir').addEventListener('click', function() {
    window.print();
});