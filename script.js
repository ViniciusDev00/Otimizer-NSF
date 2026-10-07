// Configuração inicial obrigatória da biblioteca PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

let arquivosPdf = [];
let dadosMondayExcel = null;

// Monitora o upload de PDFs (Passo 1)
document.getElementById('input-pdf').addEventListener('change', function(e) {
    arquivosPdf = Array.from(e.target.files);
    document.getElementById('lista-pdfs').innerHTML = arquivosPdf.length > 0 
        ? `✔️ ${arquivosPdf.length} PDF(s) pronto(s).` 
        : "Nenhum PDF selecionado";
    verificarRequisitos();
});

// Monitora o upload da Planilha Excel .xlsx do Monday (Passo 2)
document.getElementById('input-monday').addEventListener('change', function(e) {
    const arquivo = e.target.files[0];
    if (arquivo) {
        const leitor = new FileReader();
        leitor.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Pega a primeira aba da planilha do Excel
            const primeiraAba = workbook.SheetNames[0];
            const planilha = workbook.Sheets[primeiraAba];
            
            // Carrega em formato de matriz pura (linhas e colunas)
            dadosMondayExcel = XLSX.utils.sheet_to_json(planilha, { header: 1 });
            
            document.getElementById('status-monday').innerHTML = `✔️ Planilha "${arquivo.name}" carregada com sucesso.`;
            verificarRequisitos();
        };
        leitor.readAsArrayBuffer(arquivo);
    }
});

// Libera o botão de processamento se ambas as fontes de dados estiverem carregadas
function verificarRequisitos() {
    const btnProcessar = document.getElementById('btn-processar');
    btnProcessar.disabled = !(arquivosPdf.length > 0 && dadosMondayExcel !== null);
}

// Função de busca precisa do Aviso na Coluna F (Índice 5) do Excel
function buscarAvisoNoExcel(numeroPedido) {
    if (!dadosMondayExcel || !numeroPedido) return "Não Encontrado";
    
    let pedidoProcurado = numeroPedido.trim().toUpperCase();

    // Varre as linhas da planilha (pulando a linha 0 de cabeçalho)
    for (let i = 1; i < dadosMondayExcel.length; i++) {
        let linha = dadosMondayExcel[i];
        if (!linha || linha.length === 0) continue;

        let achouLinha = false;
        
        // Varre as células da linha corrente para verificar se ela pertence ao pedido buscado
        for (let j = 0; j < linha.length; j++) {
            let valorCelula = String(linha[j] || '').trim().toUpperCase();
            if (valorCelula === pedidoProcurado || valorCelula.includes(pedidoProcurado)) {
                achouLinha = true;
                break;
            }
        }

        if (achouLinha) {
            // Coluna F corresponde estritamente ao índice 5 (A=0, B=1, C=2, D=3, E=4, F=5)
            let valorAviso = String(linha[5] || '').trim();
            if (valorAviso) {
                // Filtra para trazer apenas os dígitos numéricos caso haja texto complementar
                let apenasNumeros = valorAviso.match(/\d+/);
                return apenasNumeros ? apenasNumeros[0] : valorAviso;
            }
        }
    }
    return "Não Encontrado";
}

// FUNÇÃO AUXILIAR: Limpa o nome do cliente tirando lixos, números de cópia e abreviações fiscais
function limparNomeCliente(nomeBruto) {
    if (!nomeBruto) return "CLIENTE DESCONHECIDO";
    
    let nome = nomeBruto.toUpperCase();
    
    // 1. Remove números de cópias de download do Windows (ex: "(1)", " (2)", " 1")
    nome = nome.replace(/\(\d+\)/g, ''); 
    nome = nome.replace(/\s\d+$/g, '');
    
    // 2. Remove códigos numéricos iniciais com hifens (ex: "001592 - ")
    nome = nome.replace(/^[0-9]+\s*-\s*/, '');
    
    // 3. Remove sufixos jurídicos e lixos textuais comuns para limpar a listagem
    nome = nome.replace(/\bLTDA\b/g, '');
    nome = nome.replace(/\bS\.A\.\b/g, '');
    nome = nome.replace(/\bSA\b/g, '');
    nome = nome.replace(/\bS\/A\b/g, '');
    nome = nome.replace(/\bME\b/g, '');
    nome = nome.replace(/\bEPP\b/g, '');
    
    // 4. Limpa aspas e espaços duplos que sobraram das remoções
    nome = nome.replace(/["']/g, '');
    nome = nome.replace(/\s+/g, ' ');
    
    return nome.trim();
}

// Processador Principal (Botão "Agrupar e Vincular Aviso")
document.getElementById('btn-processar').addEventListener('click', async function() {
    const btnImprimir = document.getElementById('btn-imprimir');
    const containerResultado = document.getElementById('resultado-impressao');
    
    // Estrutura de consolidação: { "ChaveUnica": { cliente, aviso, ordens: [], descricao, qtd } }
    const mapaAgrupadoGeral = {};

    for (let arquivo of arquivosPdf) {
        try {
            const arrayBuffer = await arquivo.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            
            // Pega o código do pedido do nome do arquivo (ex: P260422013)
            let nrPedido = "";
            let matchPedidoId = arquivo.name.match(/P\d+/i);
            if (matchPedidoId) {
                nrPedido = matchPedidoId[0].toUpperCase();
            }
            
            // Nome alternativo baseado no que sobrar do nome do arquivo limpo
            let nomeClienteAlternativo = arquivo.name
                .replace(/P\d+/i, '') // remove o código do pedido
                .replace(/\.pdf$/i, '') // remove a extensão .pdf
                .replace(/[-_()]/g, ' ') // limpa caracteres especiais comuns
                .trim();

            for (let i = 1; i <= pdf.numPages; i++) {
                const pagina = await pdf.getPage(i);
                const conteudoTexto = await pagina.getTextContent();
                const linhasTexto = conteudoTexto.items.map(item => item.str.trim()).filter(s => s !== "");

                // Se não localizou o número do pedido no nome do arquivo, busca no texto interno
                if (!nrPedido) {
                    let idxPed = linhasTexto.findIndex(t => t.toUpperCase().includes("PEDIDO:"));
                    if (idxPed !== -1 && linhasTexto[idxPed + 1]) {
                        nrPedido = linhasTexto[idxPed + 1].replace(/["']/g, '').trim().toUpperCase();
                    }
                }

                let nomeClienteBruto = "";

                // Procura por linhas com padrão de código de cliente do sistema (ex: "001592 - ")
                let linhaComCodigo = linhasTexto.find(t => /^\d+\s*-\s*/.test(t));

                if (linhaComCodigo) {
                    let nomeLimpo = linhaComCodigo.replace(/["']/g, '').trim();
                    if (!nomeLimpo.includes("- SP") && !nomeLimpo.includes("- MG") && !nomeLimpo.includes("- PE") && nomeLimpo.length > 2) {
                        nomeClienteBruto = nomeLimpo;
                    }
                }

                // Se a varredura interna falhar, utiliza o nome recuperado do título do arquivo PDF
                if (!nomeClienteBruto || nomeClienteBruto === "DESCONHECIDO") {
                    nomeClienteBruto = nomeClienteAlternativo || "CLIENTE DESCONHECIDO";
                }
                
                // Aplicação da limpeza no nome do cliente
                let nomeClienteFinal = limparNomeCliente(nomeClienteBruto);

                // Realiza o cruzamento de dados buscando o aviso correto no Excel do Monday
                let avisoVinculado = buscarAvisoNoExcel(nrPedido);

                // Varre a tabela de itens procurando linhas com formato de quantidade (ex: "36,00")
                for (let j = 0; j < linhasTexto.length; j++) {
                    if (/^\d+,\d{2}$/.test(linhasTexto[j])) {
                        let qtd = parseFloat(linhasTexto[j].replace(',', '.'));
                        
                        if (j >= 2) {
                            let descricao = linhasTexto[j - 1].replace(/["']/g, '').trim();
                            let produto = linhasTexto[j - 2].replace(/["']/g, '').trim();
                            
                            // Ignora strings de cabeçalho do relatório
                            if (produto.toUpperCase() === "PRODUTO" || produto.toUpperCase() === "DESCRIÇÃO" || produto.toUpperCase() === "LOCALIZAÇÃO") {
                                continue;
                            }

                            // Chave composta para consolidar produtos idênticos destinados ao mesmo Cliente e Aviso
                            let chaveAgrupamento = `${nomeClienteFinal}_${avisoVinculado}_${descricao}`;

                            if (mapaAgrupadoGeral[chaveAgrupamento]) {
                                mapaAgrupadoGeral[chaveAgrupamento].qtd += qtd;
                                if (nrPedido && !mapaAgrupadoGeral[chaveAgrupamento].ordens.includes(nrPedido)) {
                                    mapaAgrupadoGeral[chaveAgrupamento].ordens.push(nrPedido);
                                }
                            } else {
                                mapaAgrupadoGeral[chaveAgrupamento] = {
                                    cliente: nomeClienteFinal,
                                    aviso: avisoVinculado,
                                    ordens: nrPedido ? [nrPedido] : ["-"],
                                    descricao: descricao,
                                    qtd: qtd
                                };
                            }
                        }
                    }
                }
            }
        } catch (erro) {
            console.error("Erro crítico ao processar o arquivo: " + arquivo.name, erro);
        }
    }

    if (Object.keys(mapaAgrupadoGeral).length === 0) {
        alert("Nenhum dado pôde ser extraído. Verifique se os arquivos de Ordem estão corretos.");
        return;
    }

    // Captura as opções selecionadas na interface
    const modoRelatorio = document.querySelector('input[name="opcao-modo"]:checked').value;
    const tipoOrdenacao = document.querySelector('input[name="opcao-ordenacao"]:checked').value;
    const registrosOrdenados = Object.values(mapaAgrupadoGeral);

    // Ordena os registros
    if (tipoOrdenacao === "cliente") {
        registrosOrdenados.sort((a, b) => a.cliente.localeCompare(b.cliente) || a.descricao.localeCompare(b.descricao));
    } else if (tipoOrdenacao === "peca") {
        registrosOrdenados.sort((a, b) => a.descricao.localeCompare(b.descricao) || a.cliente.localeCompare(b.cliente));
    }

    // Aplica ou remove a classe do body para adaptar os estilos de impressão
    if (modoRelatorio === "corporativo") {
        document.body.classList.add("modo-corporativo");
    } else {
        document.body.classList.remove("modo-corporativo");
    }

    let htmlFinal = "";

    // ==========================================
    // 📊 OPÇÃO 1: MODO CORPORATIVO COMPACTO (1 PÁGINA)
    // ==========================================
    if (modoRelatorio === "corporativo") {
        let totalPecasGeral = registrosOrdenados.reduce((acc, item) => acc + item.qtd, 0);

        htmlFinal = `
            <div class="header-compacto">
                <div>
                    <h4>📊 CONTROLE GERENCIAL DE SEPARAÇÃO CONSOLIDADO</h4>
                </div>
                <div class="meta-info">
                    Emissão: ${new Date().toLocaleDateString('pt-BR')} às ${new Date().toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'})} | 
                    <strong>Total Peças: ${totalPecasGeral}</strong> | 
                    <strong>Linhas: ${registrosOrdenados.length}</strong>
                </div>
            </div>

            <table class="tabela-corporativa-compacta">
                <thead>
                    <tr>
                        <th style="width: 25%;">Cliente</th>
                        <th style="width: 12%;">N° Aviso</th>
                        <th style="width: 13%;">Ordem(ns)</th>
                        <th style="width: 42%;">Descrição do Item / Peça</th>
                        <th style="width: 8%; text-align: center;">Qtd</th>
                    </tr>
                </thead>
                <tbody>
        `;

        registrosOrdenados.forEach(reg => {
            let textoOrdens = reg.ordens.join(', ');
            htmlFinal += `
                <tr>
                    <td><strong>${reg.cliente}</strong></td>
                    <td><span style="font-family: monospace; font-weight: bold;">${reg.aviso}</span></td>
                    <td><span style="font-family: monospace;">${textoOrdens}</span></td>
                    <td>${reg.descricao}</td>
                    <td style="text-align: center; font-weight: bold;">${reg.qtd}</td>
                </tr>
            `;
        });

        htmlFinal += `
                </tbody>
            </table>
        `;

    // ==========================================
    // 📋 OPÇÃO 2: MODO PADRÃO (SEPARAÇÃO POR CLIENTE)
    // ==========================================
    } else {
        htmlFinal = `
            <div class="header-impressao">
                <div class="header-logo">
                    <img src="Logo.png" alt="Logo" onerror="this.style.display='none'">
                </div>
                <div class="header-titulo">
                    <h3>📋 PROGRAMAÇÃO DE SEPARAÇÃO CONSOLIDADA</h3>
                    <div class="meta-data">Data de Emissão: ${new Date().toLocaleDateString('pt-BR')} às ${new Date().toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'})}</div>
                </div>
            </div>
        `;

        let clienteAtual = "";
        let tabelaAberta = false;

        registrosOrdenados.forEach(reg => {
            if (reg.cliente !== clienteAtual) {
                if (tabelaAberta) {
                    htmlFinal += `</tbody></table></div>`;
                }
                
                clienteAtual = reg.cliente;
                tabelaAberta = true;

                htmlFinal += `
                    <div class="bloco-cliente-impressao" style="margin-bottom: 15px; page-break-after: avoid;">
                        <div style="background-color: #0f172a; color: #ffffff; padding: 6px 10px; font-weight: bold; font-size: 13px; border-radius: 4px 4px 0 0; text-transform: uppercase;">
                            👤 Cliente: ${clienteAtual}
                        </div>
                        <table class="tabela-separacao" style="border: 1px solid #cbd5e1; border-top: none; width: 100%;">
                            <thead>
                                <tr>
                                    <th style="width: 14%;">N° de Aviso</th>
                                    <th style="width: 14%;">N° da Ordem</th>
                                    <th style="width: 60%;">Descrição do Item</th>
                                    <th style="width: 8%; text-align: center;">Qtd Tot</th>
                                    <th style="width: 4%; text-align: center;">Conf</th>
                                </tr>
                            </thead>
                            <tbody>
                `;
            }

            let textoOrdens = reg.ordens.join(', ');
            
            htmlFinal += `
                <tr style="page-break-inside: avoid;">
                    <td><span style="font-family: monospace; font-size: 13px; font-weight: bold; color: #0f172a;">${reg.aviso}</span></td>
                    <td><span style="font-family: monospace; font-size: 12px; color: #475569;">${textoOrdens}</span></td>
                    <td>${reg.descricao}</td>
                    <td style="text-align: center; font-size: 13px;"><strong>${reg.qtd}</strong></td>
                    <td class="col-check"><div class="box-check"></div></td>
                </tr>
            `;
        });

        if (tabelaAberta) {
            htmlFinal += `</tbody></table></div>`;
        }
    }

    // Renderiza o resultado e libera o botão de impressão
    containerResultado.innerHTML = htmlFinal;
    btnImprimir.disabled = false;
});

// Impressão nativa
document.getElementById('btn-imprimir').addEventListener('click', function() {
    window.print();
});