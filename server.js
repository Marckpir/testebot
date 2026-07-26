const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const XLSX = require('xlsx');
const multer = require('multer');
const { chromium, errors } = require('playwright');

let chromeAwsLambda = null;
try {
  chromeAwsLambda = require('chrome-aws-lambda');
} catch {
  chromeAwsLambda = null;
}

async function criarBrowser({ mostrarNavegador = false } = {}) {
  const args = ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'];

  if (chromeAwsLambda) {
    const executablePath = await chromeAwsLambda.executablePath;
    const launchOptions = {
      headless: true,
      args: [...args, ...(chromeAwsLambda.args || [])],
    };
    if (executablePath) {
      launchOptions.executablePath = executablePath;
    }
    return chromium.launch(launchOptions);
  }

  return chromium.launch({
    headless: !mostrarNavegador,
    args,
  });
}

const app = express();
const RESULTADOS_CACHE = new Map();
const RESULTADOS_PAGINA_CACHE = new Map();
const APP_VERSAO = '2026-07-22-node-r2-planilha';
const APP_PORTA = Number(process.env.PORT || 5050);

// Upload da planilha (Nota / InícPlanej) fica em memória, sem gravar em disco.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
});

app.use(express.urlencoded({ extended: true }));
app.use('/static', express.static(path.join(__dirname, 'static')));

function textoLimpo(valor) {
  return String(valor || '').replace(/\s+/g, ' ').trim();
}

function escapeHtml(valor) {
  return String(valor || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizarNomeApenasTexto(valor) {
  const texto = textoLimpo(valor);
  if (!texto) {
    return '';
  }

  const somenteLetrasEspacos = [...texto]
    .map((caractere) => (/[\p{L}\s]/u.test(caractere) ? caractere : ' '))
    .join('');

  return textoLimpo(somenteLetrasEspacos);
}

function parsearListaNs(textoLista) {
  const itens = String(textoLista || '').trim().split(/[\s,;]+/);
  const vistos = new Set();
  const resultado = [];

  for (const item of itens) {
    const ns = textoLimpo(item);
    if (!ns || vistos.has(ns)) {
      continue;
    }
    vistos.add(ns);
    resultado.push(ns);
  }

  return resultado;
}

// =====================================================
// SUPORTE A PLANILHA (Nota / InícPlanej / Possui_Anexo)
// =====================================================

function normalizarChaveCabecalho(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

// Converte um número de série do Excel (sistema de datas 1900) em objeto Date.
function excelSerialParaData(serial) {
  const numero = Number(serial);
  if (!Number.isFinite(numero)) {
    return null;
  }
  const millis = Math.round((numero - 25569) * 86400 * 1000);
  const data = new Date(millis);
  return Number.isNaN(data.getTime()) ? null : data;
}

// Aceita "22/07/2026", "22/07/2026 14:32" ou "22/07/2026 14:32:10".
function parseDataBR(texto) {
  const str = textoLimpo(texto);
  const match = str.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:\D+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!match) {
    return null;
  }

  let [, dia, mes, ano, hora, minuto, segundo] = match;
  dia = Number(dia);
  mes = Number(mes);
  ano = Number(ano);
  if (ano < 100) {
    ano += 2000;
  }
  hora = hora ? Number(hora) : 0;
  minuto = minuto ? Number(minuto) : 0;
  segundo = segundo ? Number(segundo) : 0;

  const data = new Date(ano, mes - 1, dia, hora, minuto, segundo);
  return Number.isNaN(data.getTime()) ? null : data;
}

// Aceita "2026-07-22" ou "2026-07-22T14:32:00".
function parseDataISO(texto) {
  const str = textoLimpo(texto);
  const match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) {
    return null;
  }
  const [, ano, mes, dia] = match;
  const data = new Date(Number(ano), Number(mes) - 1, Number(dia));
  return Number.isNaN(data.getTime()) ? null : data;
}

// Converte qualquer valor vindo da planilha ou do APRWEB (Date, número de série, texto) em Date.
function converterValorParaData(valor) {
  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    return valor;
  }
  if (typeof valor === 'number') {
    return excelSerialParaData(valor);
  }
  const texto = textoLimpo(valor);
  if (!texto) {
    return null;
  }
  return parseDataISO(texto) || parseDataBR(texto);
}

function dataSomenteDia(data) {
  return new Date(data.getFullYear(), data.getMonth(), data.getDate());
}

function formatarDataBR(data) {
  if (!(data instanceof Date) || Number.isNaN(data.getTime())) {
    return '';
  }
  const dia = String(data.getDate()).padStart(2, '0');
  const mes = String(data.getMonth() + 1).padStart(2, '0');
  const ano = data.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

// Lê o buffer .xlsx enviado pelo usuário e localiza as colunas "Nota" e "InícPlanej"
// de forma tolerante a acentuação/maiúsculas (ex.: "Nota", "NOTA", "InícPlanej", "Início Planejado").
function lerLinhasPlanilhaNotas(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const nomeAba = workbook.SheetNames[0];
  if (!nomeAba) {
    return { linhas: [], erro: 'A planilha enviada não contém nenhuma aba.' };
  }

  const planilha = workbook.Sheets[nomeAba];
  const matriz = XLSX.utils.sheet_to_json(planilha, { header: 1, defval: '', raw: true });

  if (!matriz.length) {
    return { linhas: [], erro: 'A planilha enviada está vazia.' };
  }

  const cabecalho = matriz[0].map((valor) => normalizarChaveCabecalho(valor));

  let idxNota = cabecalho.findIndex((chave) => chave === 'nota' || chave === 'ns' || chave === 'numerons');
  if (idxNota === -1) {
    idxNota = cabecalho.findIndex((chave) => chave.startsWith('nota'));
  }

  let idxData = cabecalho.findIndex((chave) => chave.includes('inic') && chave.includes('planej'));
  if (idxData === -1) {
    idxData = cabecalho.findIndex((chave) => chave.includes('datainicio') || chave.includes('inicioplanejado'));
  }

  if (idxNota === -1) {
    idxNota = 0;
  }
  if (idxData === -1) {
    idxData = idxNota === 0 ? 1 : 0;
  }

  const linhas = [];
  for (let i = 1; i < matriz.length; i += 1) {
    const linhaRaw = matriz[i];
    if (!linhaRaw || linhaRaw.every((valor) => textoLimpo(valor) === '')) {
      continue;
    }

    const nota = textoLimpo(linhaRaw[idxNota]);
    if (!nota) {
      continue;
    }

    linhas.push({
      nota,
      valorDataPlanejada: linhaRaw[idxData],
    });
  }

  return { linhas, erro: '' };
}

// Varre o histórico de postagem de todos os anexos já extraídos (extrairAnexos/extrairHistoricoArquivo)
// e verifica se alguma data de postagem é >= à data de início planejado da nota.
function verificarAnexoPosterior(anexos, dataInicioPlanejado) {
  const dataBaseComparacao = dataInicioPlanejado ? dataSomenteDia(dataInicioPlanejado) : null;
  let dataMaisRecente = null;
  let totalDatasValidas = 0;
  let possuiPosterior = false;

  for (const anexo of anexos || []) {
    const historico = Array.isArray(anexo.historico) ? anexo.historico : [];
    for (const item of historico) {
      const dataItem = converterValorParaData(item.data);
      if (!dataItem) {
        continue;
      }

      totalDatasValidas += 1;
      if (!dataMaisRecente || dataItem > dataMaisRecente) {
        dataMaisRecente = dataItem;
      }

      if (dataBaseComparacao && dataSomenteDia(dataItem) >= dataBaseComparacao) {
        possuiPosterior = true;
      }
    }
  }

  return { possuiPosterior, dataMaisRecente, totalDatasValidas };
}

async function lerInputSeguro(page, seletor) {
  try {
    const campo = page.locator(seletor);
    if ((await campo.count()) === 0) {
      return '';
    }
    return await campo.first().inputValue();
  } catch {
    return '';
  }
}

function nomeCampoPorId(rawId) {
  if (!rawId) {
    return '';
  }

  let nome = rawId.split('$').pop();
  if (nome.includes('_')) {
    nome = nome.split('_', 2)[1] || nome;
  }

  nome = nome
    .replaceAll('textBox', '')
    .replaceAll('dropDownList', '')
    .replaceAll('label', '')
    .replaceAll('TextBox', '')
    .replaceAll('DropDownList', '')
    .replaceAll('Label', '');

  let convertido = '';
  for (let indice = 0; indice < nome.length; indice += 1) {
    const atual = nome[indice];
    const anterior = indice > 0 ? nome[indice - 1] : '';
    if (indice > 0 && /[A-Z]/.test(atual) && /[a-z]/.test(anterior)) {
      convertido += ' ';
    }
    convertido += atual;
  }

  return textoLimpo(convertido).toUpperCase();
}

async function extrairDadosDinamicos(page) {
  const dados = {};
  const seletores = [
    "input[id*='ContentPlaceHolder_']",
    "textarea[id*='ContentPlaceHolder_']",
    "select[id*='ContentPlaceHolder_']",
  ];

  for (const seletor of seletores) {
    const elementos = page.locator(seletor);
    const total = await elementos.count();

    for (let indice = 0; indice < total; indice += 1) {
      const elemento = elementos.nth(indice);

      try {
        if (!(await elemento.isVisible())) {
          continue;
        }
      } catch {
        continue;
      }

      let tipo = '';
      try {
        tipo = String((await elemento.getAttribute('type')) || '').toLowerCase();
      } catch {
        tipo = '';
      }

      if (['hidden', 'password', 'submit', 'button', 'image', 'file', 'checkbox', 'radio'].includes(tipo)) {
        continue;
      }

      const campoId = (await elemento.getAttribute('id')) || '';
      if (campoId && ['senha', 'password'].some((chave) => campoId.toLowerCase().includes(chave))) {
        continue;
      }

      let tagName = '';
      try {
        tagName = String((await elemento.evaluate((el) => el.tagName)) || '').toLowerCase();
      } catch {
        tagName = '';
      }

      let valor = '';
      if (tagName === 'select') {
        try {
          valor = textoLimpo(await elemento.evaluate((el) => el.options[el.selectedIndex]?.text || ''));
        } catch {
          valor = '';
        }
      } else {
        try {
          valor = textoLimpo(await elemento.inputValue());
        } catch {
          valor = '';
        }
      }

      if (!valor) {
        continue;
      }

      const nomeCampo = nomeCampoPorId(campoId);
      if (!nomeCampo) {
        continue;
      }

      if (!dados[nomeCampo]) {
        dados[nomeCampo] = valor;
      }
    }
  }

  return dados;
}

async function credenciaisAprInvalidas(page) {
  const termosErro = [
    'login ou senha',
    'usuário ou senha',
    'usuario ou senha',
    'senha incorreta',
    'usuário incorreto',
    'usuario incorreto',
    'acesso negado',
    'falha no login',
  ];

  const seletoresMensagem = [
    '#LabelMensagem',
    "span[id*='LabelMensagem']",
    "div[id*='validationSummary']",
    "span[id*='RequiredFieldValidator']",
    "span[id*='CustomValidator']",
    '.validation-summary-errors',
    '.erro',
    '.error',
  ];

  for (const seletor of seletoresMensagem) {
    try {
      const msg = textoLimpo(await page.locator(seletor).first().innerText());
      if (msg && termosErro.some((termo) => msg.toLowerCase().includes(termo))) {
        return true;
      }
    } catch {
      // ignora
    }
  }

  try {
    const textoPagina = textoLimpo(await page.innerText('body')).toLowerCase();
    return termosErro.some((termo) => textoPagina.includes(termo));
  } catch {
    return false;
  }
}

async function fecharModalHistorico(page) {
  const overlay = '#ContentPlaceHolder_ucHistoricoArquivo_modalPopupExtenderHistoricoArquivo_backgroundElement';
  const panel = '#ContentPlaceHolder_ucHistoricoArquivo_panelHistoricoArquivo';

  try {
    const overlayLoc = page.locator(overlay);
    if ((await overlayLoc.count()) && (await overlayLoc.first().isVisible())) {
      const seletores = [
        '#ContentPlaceHolder_ucHistoricoArquivo_imageButtonFechar',
        '#ContentPlaceHolder_ucHistoricoArquivo_linkButtonFechar',
        '#ContentPlaceHolder_ucHistoricoArquivo_buttonFechar',
      ];

      for (const seletor of seletores) {
        const botao = page.locator(seletor);
        if ((await botao.count()) && (await botao.first().isVisible())) {
          await botao.first().click({ force: true });
          break;
        }
      }

      await page.keyboard.press('Escape').catch(() => {});

      await page.waitForSelector(overlay, { state: 'hidden', timeout: 2000 }).catch(async () => {
        await page.evaluate(() => {
          const bg = document.querySelector('#ContentPlaceHolder_ucHistoricoArquivo_modalPopupExtenderHistoricoArquivo_backgroundElement');
          const pnl = document.querySelector('#ContentPlaceHolder_ucHistoricoArquivo_panelHistoricoArquivo');
          if (bg) bg.style.display = 'none';
          if (pnl) pnl.style.display = 'none';
        });
      });

      await page.waitForSelector(panel, { state: 'hidden', timeout: 2000 }).catch(() => {});
    }
  } catch {
    // ignora
  }
}

async function extrairHistoricoArquivo(page, historicoLink) {
  const historico = [];
  if ((await historicoLink.count()) === 0) {
    return historico;
  }

  await fecharModalHistorico(page);
  await historicoLink.first().click({ force: true });

  const seletoresTabelaHistorico = [
    '#ContentPlaceHolder_ucHistoricoArquivo_gridViewArquivos',
    "table[id*='gridViewArquivos']",
    '#ContentPlaceHolder_ucHistoricoArquivo_panelHistoricoArquivo table',
    "div[id*='HistoricoArquivo'] table",
  ];

  let tabelaHistorico = null;
  for (const seletor of seletoresTabelaHistorico) {
    try {
      await page.waitForSelector(seletor, { state: 'visible', timeout: 4000 });
      const candidato = page.locator(seletor);
      if ((await candidato.count()) && (await candidato.first().isVisible())) {
        tabelaHistorico = candidato.first();
        break;
      }
    } catch {
      // ignora
    }
  }

  if (!tabelaHistorico) {
    return historico;
  }

  const linhasHistorico = tabelaHistorico.locator('tr');
  const totalLinhas = await linhasHistorico.count();

  let cabecalhos = [];
  try {
    const linhaCabecalho = tabelaHistorico.locator('thead th, tr th');
    const totalCabecalhos = await linhaCabecalho.count();
    cabecalhos = [];
    for (let indice = 0; indice < totalCabecalhos; indice += 1) {
      const titulo = textoLimpo(await linhaCabecalho.nth(indice).innerText());
      cabecalhos.push(titulo || `COLUNA ${indice + 1}`);
    }
  } catch {
    cabecalhos = [];
  }

  for (let indice = 0; indice < totalLinhas; indice += 1) {
    const linha = linhasHistorico.nth(indice);
    const colunas = linha.locator('td');
    const totalColunas = await colunas.count();
    if (totalColunas === 0) {
      continue;
    }

    const valoresColunasRaw = [];
    for (let idxColuna = 0; idxColuna < totalColunas; idxColuna += 1) {
      valoresColunasRaw.push(textoLimpo(await colunas.nth(idxColuna).innerText()));
    }

    const valoresColunas = valoresColunasRaw.filter(Boolean);
    if (valoresColunas.length === 0) {
      continue;
    }

    const cabecalhoLinha = valoresColunas.join(' ').toLowerCase();
    if (cabecalhoLinha.includes('data') && (cabecalhoLinha.includes('usuário') || cabecalhoLinha.includes('usuario'))) {
      continue;
    }

    let colunasMap = {};
    if (cabecalhos.length && cabecalhos.length <= valoresColunasRaw.length) {
      for (let idx = 0; idx < cabecalhos.length; idx += 1) {
        if (valoresColunasRaw[idx]) {
          colunasMap[cabecalhos[idx]] = valoresColunasRaw[idx];
        }
      }
    } else {
      for (let idx = 0; idx < valoresColunas.length; idx += 1) {
        colunasMap[`COLUNA ${idx + 1}`] = valoresColunas[idx];
      }
    }

    let data = '';
    let usuario = '';

    Object.entries(colunasMap).forEach(([chave, valor]) => {
      const chaveNormalizada = chave.toLowerCase();
      if (!data && chaveNormalizada.includes('data')) {
        data = valor;
      }
      if (
        !usuario
        && (
          chaveNormalizada.includes('usuário')
          || chaveNormalizada.includes('usuario')
          || chaveNormalizada.includes('matrícula')
          || chaveNormalizada.includes('matricula')
          || chaveNormalizada.includes('responsável')
          || chaveNormalizada.includes('responsavel')
        )
      ) {
        usuario = valor;
      }
    });

    if (!data) {
      const candidatoData = valoresColunas.find((valor) => /\b\d{2}\/\d{2}\/\d{4}\b/.test(valor));
      if (candidatoData) {
        data = candidatoData;
      }
    }

    if (!data && valoresColunas.length) {
      data = valoresColunas[0];
    }
    if (!usuario && valoresColunas.length > 1) {
      usuario = valoresColunas[1];
    }

    historico.push({
      data,
      usuario,
      texto: valoresColunas.join(' | '),
      colunas: colunasMap,
    });
  }

  await fecharModalHistorico(page);
  return historico;
}

async function extrairAnexos(page) {
  const anexos = [];
  const chavesVistas = new Set();
  const acoesPaginacaoVistas = new Set();

  async function extrairLinhasPaginaAtual(paginaAtual) {
    let linhas = page.locator('#tableFileUpload tr');
    if ((await linhas.count()) === 0) {
      linhas = page.locator("table[id*='tableFileUpload'] tr");
    }
    if ((await linhas.count()) === 0) {
      linhas = page.locator("table[id*='FileUpload'] tr");
    }
    if ((await linhas.count()) === 0) {
      linhas = page.locator('#tableFileUpload tbody tr');
    }
    if ((await linhas.count()) === 0) {
      linhas = page.locator("tr:has(a[id*='linkButtonDownloadAnexo']), tr:has(select[id*='dropDownListSituacaoAnexo']), tr:has(span[id*='textoTipoAnexo'])");
    }

    const totalLinhas = await linhas.count();

    for (let indice = 0; indice < totalLinhas; indice += 1) {
      const linha = linhas.nth(indice);
      const colunas = linha.locator('td, th');
      const totalColunas = await colunas.count();
      if (totalColunas === 0) {
        continue;
      }

      let descricao = totalColunas >= 1 ? textoLimpo(await colunas.nth(0).innerText()) : '';

      const valoresLinha = [];
      for (let idxColuna = 0; idxColuna < totalColunas; idxColuna += 1) {
        const valorColuna = textoLimpo(await colunas.nth(idxColuna).innerText());
        if (valorColuna) {
          valoresLinha.push(valorColuna);
        }
      }

      const segundaColuna = totalColunas >= 2 ? colunas.nth(1) : colunas.nth(0);

      const inputFile = segundaColuna.locator("input[type='file']");
      const inputFileId = (await inputFile.count()) ? (await inputFile.first().getAttribute('id')) || '' : '';
      const inputFileName = (await inputFile.count()) ? (await inputFile.first().getAttribute('name')) || '' : '';

      const arquivoLink = linha.locator("a[id*='linkButtonDownloadAnexo'], a[id*='DownloadAnexo'], a[title*='Download']");
      const arquivoNome = (await arquivoLink.count()) ? textoLimpo(await arquivoLink.first().innerText()) : '(sem arquivo)';

      let downloadId = '';
      let codanexo = '';
      let nomearquivo = '';
      let numerons = '';
      let fileclassification = '';
      let ficha = '';

      if ((await arquivoLink.count())) {
        downloadId = (await arquivoLink.first().getAttribute('id')) || '';
        codanexo = (await arquivoLink.first().getAttribute('codanexo')) || '';
        nomearquivo = (await arquivoLink.first().getAttribute('nomearquivo')) || '';
        numerons = (await arquivoLink.first().getAttribute('numerons')) || '';
        fileclassification = (await arquivoLink.first().getAttribute('fileclassification')) || '';
        ficha = (await arquivoLink.first().getAttribute('ficha')) || '';
      }

      const selectSituacao = linha.locator("select[id*='dropDownListSituacaoAnexo'], select[id*='SituacaoAnexo'], select[id*='dropDownListSituacao']");
      let situacao = '(não informado)';
      let situacaoSelectId = '';
      let situacaoTitle = '';

      if ((await selectSituacao.count())) {
        situacao = textoLimpo(await selectSituacao.first().evaluate((el) => el.options[el.selectedIndex]?.text || '')) || '(não informado)';
        situacaoSelectId = (await selectSituacao.first().getAttribute('id')) || '';
        situacaoTitle = (await selectSituacao.first().getAttribute('title')) || '';
      }

      const formatoSpan = linha.locator("span[id*='textoTipoAnexo'], span[id*='TipoAnexo']");
      const formato = (await formatoSpan.count()) ? textoLimpo(await formatoSpan.first().innerText()) : '';

      if (!descricao) {
        descricao = textoLimpo(formato) || textoLimpo(arquivoNome);
      }
      if (!descricao && valoresLinha.length) {
        descricao = valoresLinha[0];
      }
      if (!descricao) {
        continue;
      }

      const historicoLink = linha.locator("a:has-text('(Histórico)'), a:has-text('Histórico'), a:has-text('Historico'), a[title*='Histórico'], a[title*='Historico'], a[id*='Historico']");
      const historicoId = (await historicoLink.count()) ? ((await historicoLink.first().getAttribute('id')) || '') : '';

      const situacaoImg = segundaColuna.locator("img[id*='imageSituacaoAnexo']");
      const situacaoImgId = (await situacaoImg.count()) ? ((await situacaoImg.first().getAttribute('id')) || '') : '';
      const situacaoImgSrc = (await situacaoImg.count()) ? ((await situacaoImg.first().getAttribute('src')) || '') : '';

      const chaveEstrutural = [codanexo, nomearquivo, inputFileName, downloadId].join('|');
      const chaveAnexo = chaveEstrutural.replaceAll('|', '').trim()
        ? [chaveEstrutural, descricao, arquivoNome].join('|')
        : [`PAG${paginaAtual}`, `LIN${indice}`, descricao, arquivoNome, valoresLinha.join(' | ')].join('|');

      if (chavesVistas.has(chaveAnexo)) {
        continue;
      }

      let historico = [];
      try {
        historico = await extrairHistoricoArquivo(page, historicoLink);
      } catch {
        historico = [];
      }

      chavesVistas.add(chaveAnexo);
      anexos.push({
        descricao,
        arquivo: arquivoNome,
        situacao,
        formato,
        input_file_id: inputFileId,
        input_file_name: inputFileName,
        download_id: downloadId,
        codanexo,
        nomearquivo,
        numerons,
        fileclassification,
        ficha,
        historico_id: historicoId,
        situacao_select_id: situacaoSelectId,
        situacao_title: situacaoTitle,
        situacao_img_id: situacaoImgId,
        situacao_img_src: situacaoImgSrc,
        linha_texto: valoresLinha.join(' | '),
        historico,
      });
    }
  }

  async function clicarProximaPagina() {
    let assinaturaAntes = '';
    try {
      assinaturaAntes = textoLimpo(await page.locator('#tableFileUpload').first().innerText());
    } catch {
      try {
        assinaturaAntes = textoLimpo(await page.locator("table[id*='FileUpload']").first().innerText());
      } catch {
        assinaturaAntes = '';
      }
    }

    let linksPaginacao = page.locator("#tableFileUpload a[href*='Page$']");
    if ((await linksPaginacao.count()) === 0) {
      linksPaginacao = page.locator("table[id*='FileUpload'] a[href*='Page$']");
    }

    const totalLinks = await linksPaginacao.count();

    for (let idx = 0; idx < totalLinks; idx += 1) {
      const link = linksPaginacao.nth(idx);
      try {
        if (!(await link.isVisible())) {
          continue;
        }
      } catch {
        continue;
      }

      const href = ((await link.getAttribute('href')) || '').trim();
      const textoLink = textoLimpo(await link.innerText());
      const acao = href || textoLink;

      if (!acao || acoesPaginacaoVistas.has(acao)) {
        continue;
      }

      if (!acao.includes('Page$') && !['>', '»', 'Próx', 'Prox'].some((ch) => textoLink.includes(ch))) {
        continue;
      }

      acoesPaginacaoVistas.add(acao);

      try {
        await link.click({ force: true });
        await page.waitForLoadState('networkidle');
        await page.waitForTimeout(700);
      } catch {
        continue;
      }

      let assinaturaDepois = '';
      try {
        assinaturaDepois = textoLimpo(await page.locator('#tableFileUpload').first().innerText());
      } catch {
        try {
          assinaturaDepois = textoLimpo(await page.locator("table[id*='FileUpload']").first().innerText());
        } catch {
          assinaturaDepois = '';
        }
      }

      if (assinaturaDepois && assinaturaDepois !== assinaturaAntes) {
        return true;
      }
    }

    const seletoresProximo = [
      "#tableFileUpload a[href*='Page$Next']",
      "table[id*='FileUpload'] a[href*='Page$Next']",
      "#tableFileUpload a[title*='Próx']",
      "#tableFileUpload a[title*='Prox']",
      "#tableFileUpload a:has-text('Próximo')",
      "#tableFileUpload a:has-text('Proximo')",
      "table[id*='FileUpload'] a[title*='Próx']",
      "table[id*='FileUpload'] a[title*='Prox']",
      "table[id*='FileUpload'] a:has-text('Próximo')",
      "table[id*='FileUpload'] a:has-text('Proximo')",
      "#tableFileUpload a:has-text('>')",
      "#tableFileUpload a:has-text('»')",
      "table[id*='FileUpload'] a:has-text('>')",
      "table[id*='FileUpload'] a:has-text('»')",
    ];

    for (const seletor of seletoresProximo) {
      const link = page.locator(seletor);
      if ((await link.count()) && (await link.first().isVisible())) {
        try {
          const href = ((await link.first().getAttribute('href')) || '').trim();
          const textoLink = textoLimpo(await link.first().innerText());
          const acao = href || `fallback::${textoLink}`;
          if (acoesPaginacaoVistas.has(acao)) {
            continue;
          }
          acoesPaginacaoVistas.add(acao);
        } catch {
          // ignora
        }

        try {
          await link.first().click({ force: true });
          await page.waitForLoadState('networkidle');
          await page.waitForTimeout(700);
        } catch {
          continue;
        }

        let assinaturaDepois = '';
        try {
          assinaturaDepois = textoLimpo(await page.locator('#tableFileUpload').first().innerText());
        } catch {
          try {
            assinaturaDepois = textoLimpo(await page.locator("table[id*='FileUpload']").first().innerText());
          } catch {
            assinaturaDepois = '';
          }
        }

        if (assinaturaDepois && assinaturaDepois !== assinaturaAntes) {
          return true;
        }
      }
    }

    return false;
  }

  const maxPaginas = 12;
  for (let paginaAtual = 1; paginaAtual <= maxPaginas; paginaAtual += 1) {
    await extrairLinhasPaginaAtual(paginaAtual);
    const teveProxima = await clicarProximaPagina();
    if (!teveProxima) {
      break;
    }
  }

  return anexos;
}

async function garantirSecaoAnexosVisivel(page) {
  const candidatos = [
    "a:has-text('Anexos')",
    "button:has-text('Anexos')",
    "[id*='Anexos']",
  ];

  const tabela = page.locator("#tableFileUpload, table[id*='tableFileUpload'], table[id*='FileUpload']");
  if ((await tabela.count()) && (await tabela.first().isVisible())) {
    return;
  }

  for (const seletor of candidatos) {
    const alvo = page.locator(seletor);
    if ((await alvo.count()) && (await alvo.first().isVisible())) {
      try {
        await alvo.first().click({ force: true });
        await page.waitForTimeout(600);
        if ((await tabela.count()) && (await tabela.first().isVisible())) {
          return;
        }
      } catch {
        // ignora
      }
    }
  }
}

async function aguardarAnexosRenderizados(page) {
  try {
    await page.waitForFunction(() => {
      return document.querySelectorAll("a[id*='linkButtonDownloadAnexo']").length > 0
        || document.querySelectorAll('#tableFileUpload tr').length > 1
        || document.querySelectorAll("table[id*='FileUpload'] tr").length > 1
        || document.querySelectorAll("select[id*='dropDownListSituacaoAnexo']").length > 0;
    }, { timeout: 15000 });
  } catch {
    // ignora
  }
}

// Erro específico para quando a busca da NS não abre o registro (botão "Visualizar" não aparece).
// É tratado separadamente de falhas de login/rede para poder ser classificado como "ausência de anexo"
// em vez de "Erro" no relatório de planilha.
class NsSemResultadoError extends Error {
  constructor(mensagem) {
    super(mensagem);
    this.name = 'NsSemResultadoError';
  }
}

// Faz login no APRWEB, pesquisa a NS informada e abre o registro (clique em "Visualizar").
// Reaproveitado tanto pela consulta normal quanto pelo novo modo de download de documentos.
async function abrirNotaNoApr(page, numeroNs, usuarioApr, senhaApr) {
  await page.goto('https://partapr.cemig.com.br');

  await page.waitForSelector('#linkButtonApr', { state: 'visible' });
  await page.click('#linkButtonApr');

  await page.waitForSelector('#TextBoxMatricula', { state: 'visible' });
  await page.fill('#TextBoxMatricula', usuarioApr);
  await page.fill('#TextBoxSenha', senhaApr);

  await page.waitForSelector('#ButtonAcessar', { state: 'visible' });
  await page.click('#ButtonAcessar');
  await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});

  try {
    await page.waitForSelector('a[title="Pedidos APR"]', { state: 'visible', timeout: 12000 });
  } catch {
    try {
      if ((await page.locator('#TextBoxMatricula').count()) && (await page.locator('#TextBoxMatricula').first().isVisible())) {
        throw new Error('Login ou senha incorretos.');
      }
    } catch (erroLogin) {
      if (erroLogin.message.includes('Login ou senha incorretos')) {
        throw erroLogin;
      }
    }

    if (await credenciaisAprInvalidas(page)) {
      throw new Error('Login ou senha incorretos.');
    }
    throw new Error('Não foi possível autenticar no APR. Verifique login, senha e tente novamente.');
  }

  await page.click('a[title="Pedidos APR"]');
  await page.waitForLoadState('networkidle');

  await page.waitForSelector('#ContentPlaceHolder_textBoxNumeroNotaServico', { state: 'visible' });
  await page.fill('#ContentPlaceHolder_textBoxNumeroNotaServico', numeroNs);

  await page.waitForSelector('#ContentPlaceHolder_imageButtonPesquisar', { state: 'visible' });
  await page.click('#ContentPlaceHolder_imageButtonPesquisar');
  await page.waitForLoadState('networkidle');

  await page.waitForSelector('input[type="image"][alt="Visualizar"]', { state: 'visible', timeout: 15000 })
    .catch(() => {});

  const botaoVisualizar = page.locator('input[type="image"][alt="Visualizar"]');
  if ((await botaoVisualizar.count()) === 0) {
    throw new NsSemResultadoError(
      `Nenhum registro encontrado para a NS ${numeroNs} no APR (nota não localizada ou sem retorno na pesquisa).`
    );
  }

  await botaoVisualizar.first().click();
  await page.waitForLoadState('networkidle');

  await page.waitForSelector('#ContentPlaceHolder_textBoxResponsavelCadastro', { state: 'visible' });
}

async function consultarNsApr(numeroNs, usuarioApr, senhaApr, { mostrarNavegador = false, incluirAnexos = true } = {}) {
  const browser = await criarBrowser({ mostrarNavegador });

  try {
    const page = await browser.newPage();
    await abrirNotaNoApr(page, numeroNs, usuarioApr, senhaApr);

    const dadosBase = {
      'NOME RESPONSAVEL TECNICO': await lerInputSeguro(page, '#ContentPlaceHolder_textBoxResponsavelCadastro'),
      'NOME DO CLIENTE': await lerInputSeguro(page, '#ContentPlaceHolder_textBoxNomeCliente'),
      TELEFONE: await lerInputSeguro(page, '#ContentPlaceHolder_textBoxTelefone'),
      EMAIL: await lerInputSeguro(page, '#ContentPlaceHolder_textBoxEmailCliente'),
      'NUMERO DE INSTALAÇÃO': await lerInputSeguro(page, '#ContentPlaceHolder_textBoxNumeroInstalacao'),
      LOGRADOURO: await lerInputSeguro(page, '#ContentPlaceHolder_textBoxLogradouro'),
      NUMERO: await lerInputSeguro(page, '#ContentPlaceHolder_textBoxNumeroEndereco'),
      BAIRRO: await lerInputSeguro(page, '#ContentPlaceHolder_textBoxBairro'),
      MUNICIPIO: await lerInputSeguro(page, '#ContentPlaceHolder_textBoxMunicipio'),
    };

    const dados = Object.fromEntries(Object.entries(dadosBase).filter(([, valor]) => textoLimpo(valor)));
    const dadosDinamicos = await extrairDadosDinamicos(page);

    for (const [chave, valor] of Object.entries(dadosDinamicos)) {
      if (!dados[chave]) {
        dados[chave] = valor;
      }
    }

    if (!incluirAnexos) {
      return { dados, anexos: [] };
    }

    await garantirSecaoAnexosVisivel(page);
    await aguardarAnexosRenderizados(page);

    await page.waitForSelector("#tableFileUpload, table[id*='FileUpload'], #tableFileUpload tbody tr, table[id*='FileUpload'] tbody tr, tr:has(a[id*='linkButtonDownloadAnexo']), tr:has(a[id*='DownloadAnexo']), tr:has(select[id*='dropDownListSituacaoAnexo']), tr:has(select[id*='SituacaoAnexo'])", { state: 'visible', timeout: 15000 }).catch(() => {});

    let anexos = await extrairAnexos(page);

    if (!anexos.length) {
      await page.waitForTimeout(2000);
      try {
        await page.reload({ waitUntil: 'networkidle' });
        await page.waitForSelector('a[title="Pedidos APR"]', { state: 'visible', timeout: 15000 });
        await page.click('a[title="Pedidos APR"]');
        await page.waitForLoadState('networkidle');

        await page.waitForSelector('#ContentPlaceHolder_textBoxNumeroNotaServico', { state: 'visible', timeout: 15000 });
        await page.fill('#ContentPlaceHolder_textBoxNumeroNotaServico', numeroNs);
        await page.click('#ContentPlaceHolder_imageButtonPesquisar');
        await page.waitForLoadState('networkidle');

        await page.waitForSelector('input[type="image"][alt="Visualizar"]', { state: 'visible', timeout: 15000 });
        await page.locator('input[type="image"][alt="Visualizar"]').first().click();
        await page.waitForLoadState('networkidle');
        await page.waitForTimeout(1500);

        await garantirSecaoAnexosVisivel(page);
        await aguardarAnexosRenderizados(page);
      } catch {
        // ignora
      }

      anexos = await extrairAnexos(page);
    }

    return { dados, anexos };
  } finally {
    await browser.close();
  }
}

// =====================================================
// DOWNLOAD DE DOCUMENTOS POSTERIORES AO INÍCIO PLANEJADO
// =====================================================

// Gera o nome da pasta no padrão NS_XXXX_DOCUMENTACAO_DDMMAAAA_HORAMINUTOSEGUNDO.
function montarNomePastaDownload(numeroNs) {
  const agora = new Date();
  const dia = String(agora.getDate()).padStart(2, '0');
  const mes = String(agora.getMonth() + 1).padStart(2, '0');
  const ano = agora.getFullYear();
  const hora = String(agora.getHours()).padStart(2, '0');
  const minuto = String(agora.getMinutes()).padStart(2, '0');
  const segundo = String(agora.getSeconds()).padStart(2, '0');
  const notaSanitizada = String(numeroNs || '').trim().replace(/[^a-zA-Z0-9_-]/g, '') || 'SEMNUMERO';

  return `NS_${notaSanitizada}_DOCUMENTACAO_${dia}${mes}${ano}_${hora}${minuto}${segundo}`;
}

function sanitizarNomeArquivo(nome) {
  const limpo = String(nome || '').trim().replace(/[\\/:*?"<>|]/g, '_');
  return limpo || 'arquivo_sem_nome';
}

// Evita sobrescrever arquivos com o mesmo nome, adicionando "(1)", "(2)", etc.
function gerarCaminhoArquivoUnico(pastaDestino, nomeArquivoSugerido) {
  const nomeSanitizado = sanitizarNomeArquivo(nomeArquivoSugerido);
  const extensao = path.extname(nomeSanitizado);
  const baseSemExtensao = path.basename(nomeSanitizado, extensao);

  let caminhoFinal = path.join(pastaDestino, nomeSanitizado);
  let contador = 1;

  while (fs.existsSync(caminhoFinal)) {
    caminhoFinal = path.join(pastaDestino, `${baseSemExtensao} (${contador})${extensao}`);
    contador += 1;
  }

  return caminhoFinal;
}

// Varre as páginas de anexos da nota já aberta e baixa apenas os arquivos cuja data de postagem
// (extraída do histórico de cada anexo) seja igual ou posterior à data de início planejado.
function extrairNomeDoContentDisposition(headerValor) {
  const texto = String(headerValor || '');
  const matchUtf8 = texto.match(/filename\*=(?:UTF-8'')?"?([^;"]+)"?/i);
  if (matchUtf8 && matchUtf8[1]) {
    try {
      return decodeURIComponent(matchUtf8[1]);
    } catch {
      return matchUtf8[1];
    }
  }
  const matchSimples = texto.match(/filename="?([^;"]+)"?/i);
  return matchSimples && matchSimples[1] ? matchSimples[1] : '';
}

// Clica no link de download e tenta capturar o arquivo por três caminhos possíveis, já que nem
// todo sistema serve o anexo com "Content-Disposition: attachment" (o que faria o Chrome baixar
// automaticamente):
//   1) evento nativo de download do navegador (caso clássico de attachment);
//   2) resposta HTTP não-HTML carregada como navegação (ex.: PDF/arquivo aberto "inline");
//   3) abertura do arquivo em uma nova aba/popup.
// Isso evita que o download falhe silenciosamente só porque o APRWEB não força o attachment.
// Instrui o Chrome, via CDP, a salvar automaticamente qualquer download nesta pasta — isso
// funciona mesmo quando o evento "download" do Playwright não é disparado (visto em execuções
// reais com o executável de Chrome instalado no Windows). Tenta o comando novo (Browser.*) e o
// antigo (Page.*), pois a disponibilidade varia entre versões do Chrome.
async function habilitarDownloadViaCdp(page, pastaDestino) {
  try {
    const client = await page.context().newCDPSession(page);
    try {
      await client.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: pastaDestino });
      return true;
    } catch {
      await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: pastaDestino });
      return true;
    }
  } catch {
    // Se nenhum dos dois comandos funcionar, seguimos apenas com a estratégia baseada em eventos.
    return false;
  }
}

function listarArquivosPasta(pastaDestino) {
  try {
    return new Set(fs.readdirSync(pastaDestino));
  } catch {
    return new Set();
  }
}

// Fica de olho na pasta de destino esperando um arquivo novo aparecer (e parar de crescer, ou
// seja, o download realmente terminou) — usado como estratégia principal já que o Chrome está
// salvando os downloads diretamente ali via CDP, sem passar pelos eventos do Playwright.
async function aguardarNovoArquivoNaPasta(pastaDestino, arquivosAntes, timeoutMs) {
  const inicio = Date.now();

  while (Date.now() - inicio < timeoutMs) {
    let atuais = [];
    try {
      atuais = fs.readdirSync(pastaDestino);
    } catch {
      atuais = [];
    }

    const novos = atuais.filter((nome) => !arquivosAntes.has(nome) && !nome.startsWith('_debug_'));
    const finalizados = novos.filter((nome) => !nome.endsWith('.crdownload') && !nome.endsWith('.tmp'));

    if (finalizados.length) {
      const caminho = path.join(pastaDestino, finalizados[0]);
      let tamanhoAnterior = -1;
      try {
        tamanhoAnterior = fs.statSync(caminho).size;
      } catch {
        tamanhoAnterior = -1;
      }

      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 500));

      let tamanhoAtual = -2;
      try {
        tamanhoAtual = fs.statSync(caminho).size;
      } catch {
        tamanhoAtual = -2;
      }

      if (tamanhoAtual >= 0 && tamanhoAtual === tamanhoAnterior) {
        return caminho;
      }
    }

    // eslint-disable-next-line no-await-in-loop
    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  return null;
}

async function baixarArquivoPorClique(page, arquivoLink, pastaDestino, nomeArquivoFallback) {
  const context = page.context();
  let downloadCapturado = null;
  const candidatosResposta = [];
  const todasRespostasObservadas = [];
  let novaPagina = null;
  let ultimoRequestCapturado = null;

  const aoBaixar = (download) => {
    // O evento de download é sempre a fonte mais confiável — tem prioridade sobre qualquer
    // resposta HTTP capturada, mesmo que ela tenha chegado primeiro.
    downloadCapturado = download;
  };

  const aoResponder = (response) => {
    try {
      const headers = response.headers();
      const contentType = (headers['content-type'] || '').toLowerCase();
      const contentDisposition = (headers['content-disposition'] || '').toLowerCase();
      const resourceType = response.request().resourceType();
      const ehDocumentoPrincipal = resourceType === 'document';

      // Guarda um resumo de TODA resposta observada na janela do clique, mesmo que não pareça
      // arquivo — vira material de diagnóstico caso nada seja capturado (ver salvarDiagnosticoFalha).
      todasRespostasObservadas.push(
        `[${response.status()}] ${resourceType} content-type="${contentType}" content-disposition="${contentDisposition}" url=${response.url()}`
      );

      const pareceArquivo = contentDisposition.includes('attachment')
        || (ehDocumentoPrincipal && contentType && !contentType.includes('text/html'));

      if (pareceArquivo && response.ok()) {
        candidatosResposta.push(response);
      }
    } catch {
      // ignora
    }
  };

  const aoAbrirPagina = (pagina) => {
    if (!novaPagina) {
      novaPagina = pagina;
    }
  };

  const cdpCandidateRequestIds = new Map();
  let cdpClient = null;

  try {
    cdpClient = await page.context().newCDPSession(page);
    await cdpClient.send('Network.enable');
    cdpClient.on('Network.responseReceived', (event) => {
      try {
        const headers = event.response.headers || {};
        const contentDisposition = (headers['content-disposition'] || '').toLowerCase();
        const contentType = (headers['content-type'] || '').toLowerCase();
        const ehDocumento = event.type === 'Document' || contentType.includes('application/pdf') || contentDisposition.includes('attachment');
        if (!ehDocumento) {
          return;
        }

        const chave = `${event.response.url}|${event.response.status}`;
        const atual = cdpCandidateRequestIds.get(chave) || [];
        atual.push({ requestId: event.requestId, response: event.response });
        cdpCandidateRequestIds.set(chave, atual);
      } catch {
        // ignora
      }
    });
  } catch {
    cdpClient = null;
  }

  context.on('download', aoBaixar);
  context.on('response', aoResponder);
  context.on('page', aoAbrirPagina);

  async function obterBufferViaCdp(response) {
    if (!cdpClient) {
      return null;
    }

    const chave = `${response.url()}|${response.status()}`;
    const candidatos = cdpCandidateRequestIds.get(chave) || [];
    for (const item of candidatos) {
      try {
        const base64 = await cdpClient.send('Network.getResponseBody', {
          requestId: item.requestId,
        });
        if (base64 && base64.body) {
          return base64.base64Encoded ? Buffer.from(base64.body, 'base64') : Buffer.from(base64.body, 'utf8');
        }
      } catch {
        // ignora e tenta próximo candidato
      }
    }
    return null;
  }

  // Salva um screenshot e um log de texto com as respostas observadas, para diagnóstico manual
  // quando nenhuma estratégia de captura funcionar.
  async function salvarDiagnosticoFalha(infoElementoClicado) {
    const timestamp = Date.now();
    const prefixo = `_debug_falha_${timestamp}`;
    try {
      await page.screenshot({ path: path.join(pastaDestino, `${prefixo}.png`) });
    } catch {
      // ignora — diagnostico é best-effort
    }
    try {
      const linhas = [
        `Elemento clicado: ${infoElementoClicado}`,
        '',
        todasRespostasObservadas.length
          ? todasRespostasObservadas.join('\n')
          : '(nenhuma resposta de rede observada durante a janela do clique)',
      ];
      fs.writeFileSync(path.join(pastaDestino, `${prefixo}.txt`), linhas.join('\n'), 'utf8');
    } catch {
      // ignora
    }
    return prefixo;
  }

  async function baixarViaFetchManual(infoRequest, nomeArquivoFallback) {
    if (!infoRequest || !infoRequest.url) {
      return null;
    }

    try {
      const resultado = await page.evaluate(async (requestInfo) => {
        const headers = {};
        for (const [key, value] of Object.entries(requestInfo.headers || {})) {
          const lower = key.toLowerCase();
          if (['host', 'content-length', 'accept-encoding', 'connection'].includes(lower)) {
            continue;
          }
          headers[key] = value;
        }

        const response = await fetch(requestInfo.url, {
          method: requestInfo.method || 'GET',
          headers,
          body: requestInfo.postData || undefined,
          credentials: 'same-origin',
        });

        if (!response.ok) {
          return { error: `status ${response.status}` };
        }

        const arrayBuffer = await response.arrayBuffer();
        const bytes = new Uint8Array(arrayBuffer);
        let binary = '';
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
          const chunk = bytes.subarray(i, i + chunkSize);
          binary += String.fromCharCode.apply(null, Array.from(chunk));
        }
        const base64 = btoa(binary);
        return {
          base64,
          contentDisposition: response.headers.get('content-disposition') || '',
        };
      }, infoRequest);

      if (!resultado || !resultado.base64) {
        return null;
      }

      const buffer = Buffer.from(resultado.base64, 'base64');
      const nomeSugerido = extrairNomeDoContentDisposition(resultado.contentDisposition) || nomeArquivoFallback;
      const caminhoFinal = gerarCaminhoArquivoUnico(pastaDestino, nomeSugerido);
      fs.writeFileSync(caminhoFinal, buffer);
      return caminhoFinal;
    } catch (error) {
      todasRespostasObservadas.push(`FALHA FETCH MANUAL: ${error && error.message ? error.message : 'erro desconhecido'}`);
      return null;
    }
  }

  let infoElementoClicado = '(não capturado)';
  const arquivosAntesClique = listarArquivosPasta(pastaDestino);

  try {
    try {
      const elemento = arquivoLink.first();
      const idAttr = (await elemento.getAttribute('id')) || '';
      const href = (await elemento.getAttribute('href')) || '';
      const onclick = (await elemento.getAttribute('onclick')) || '';
      infoElementoClicado = `id="${idAttr}" href="${href}" onclick="${onclick}"`;
    } catch {
      infoElementoClicado = '(não foi possível ler atributos do elemento antes do clique)';
    }

    const respostaDownloadProm = page.waitForResponse(
      (response) => {
        try {
          const headers = response.headers();
          const contentType = (headers['content-type'] || '').toLowerCase();
          const contentDisposition = (headers['content-disposition'] || '').toLowerCase();
          const resourceType = response.request().resourceType();
          return (
            response.ok() &&
            (contentDisposition.includes('attachment') ||
              (resourceType === 'document' && contentType && !contentType.includes('text/html')))
          );
        } catch {
          return false;
        }
      },
      { timeout: 8000 }
    ).catch(() => null);

    const downloadEventProm = context.waitForEvent('download', { timeout: 15000 }).catch(() => null);
    const navegacaoProm = page.waitForNavigation({ waitUntil: 'load', timeout: 7000 }).catch(() => null);

    const clickProm = arquivoLink.first().click({ force: true });
    const [downloadEvent] = await Promise.all([downloadEventProm, navegacaoProm, clickProm]);
    if (downloadEvent) {
      downloadCapturado = downloadEvent;
    }

    const respostaDownload = await respostaDownloadProm;
    if (respostaDownload) {
      candidatosResposta.push(respostaDownload);
    }

    // Estratégia 1 (mais confiável, conforme diagnóstico já coletado): o Chrome grava o arquivo
    // direto na pasta de destino via CDP (ver habilitarDownloadViaCdp) — só falta esperar aparecer.
    const caminhoPastaCdp = await aguardarNovoArquivoNaPasta(pastaDestino, arquivosAntesClique, 10000);
    if (caminhoPastaCdp) {
      return caminhoPastaCdp;
    }

    // Estratégia 2 (fallback): evento nativo de download do Playwright ou resposta HTTP com o arquivo.
    const prazoTotalMs = 8000;
    const margemAposCandidatoMs = 1500;
    const inicio = Date.now();
    let candidatoDetectadoEm = null;

    while (Date.now() - inicio < prazoTotalMs && !downloadCapturado) {
      if (candidatosResposta.length && !candidatoDetectadoEm) {
        candidatoDetectadoEm = Date.now();
      }
      if (candidatoDetectadoEm && Date.now() - candidatoDetectadoEm > margemAposCandidatoMs) {
        break;
      }
      // eslint-disable-next-line no-await-in-loop
      await page.waitForTimeout(200);
    }

    if (downloadCapturado) {
      const nomeSugerido = downloadCapturado.suggestedFilename() || nomeArquivoFallback;
      const caminhoFinal = gerarCaminhoArquivoUnico(pastaDestino, nomeSugerido);
      await downloadCapturado.saveAs(caminhoFinal);
      return caminhoFinal;
    }

    for (const resposta of candidatosResposta) {
      let buffer = null;
      let nomeSugerido = '';
      try {
        // eslint-disable-next-line no-await-in-loop
        buffer = await resposta.body();
        const headers = resposta.headers();
        nomeSugerido = extrairNomeDoContentDisposition(headers['content-disposition']) || nomeArquivoFallback;
      } catch (error) {
        todasRespostasObservadas.push(`FALHA LENDO CORPO: ${error && error.message ? error.message : 'erro desconhecido'}`);
      }

      if (!buffer) {
        buffer = await obterBufferViaCdp(resposta);
        try {
          const headers = resposta.headers();
          nomeSugerido = extrairNomeDoContentDisposition(headers['content-disposition']) || nomeArquivoFallback;
        } catch {
          nomeSugerido = nomeArquivoFallback;
        }
      }

      if (!buffer) {
        try {
          const request = resposta.request();
          const requestInfo = {
            url: request.url(),
            method: request.method(),
            headers: request.headers(),
            postData: request.postData(),
          };
          const caminhoFetch = await baixarViaFetchManual(requestInfo, nomeArquivoFallback);
          if (caminhoFetch) {
            return caminhoFetch;
          }
        } catch (error) {
          todasRespostasObservadas.push(`FALHA FETCH VIA REQUEST: ${error && error.message ? error.message : 'erro desconhecido'}`);
        }
      }

      if (buffer) {
        const caminhoFinal = gerarCaminhoArquivoUnico(pastaDestino, nomeSugerido || nomeArquivoFallback);
        fs.writeFileSync(caminhoFinal, buffer);
        return caminhoFinal;
      }
    }

    if (!downloadCapturado && !candidatosResposta.length && ultimoRequestCapturado) {
      const caminhoFetch = await baixarViaFetchManual(ultimoRequestCapturado, nomeArquivoFallback);
      if (caminhoFetch) {
        return caminhoFetch;
      }
    }

    // Estratégia 3 (última chance): o arquivo pode simplesmente ter demorado mais para ser
    // finalizado pelo Chrome do que a primeira janela de espera considerou.
    const caminhoPastaCdpFinal = await aguardarNovoArquivoNaPasta(pastaDestino, arquivosAntesClique, 3000);
    if (caminhoPastaCdpFinal) {
      return caminhoPastaCdpFinal;
    }

    const prefixoDiagnostico = await salvarDiagnosticoFalha(infoElementoClicado);
    throw new Error(
      `Não foi possível identificar o arquivo baixado (nem na pasta de destino, nem via evento de download ou resposta de arquivo) dentro do tempo limite. Diagnóstico salvo em ${prefixoDiagnostico}.png/.txt na pasta de destino.`
    );
  } finally {
    context.off('download', aoBaixar);
    context.off('response', aoResponder);
    context.off('page', aoAbrirPagina);

    if (cdpClient) {
      await cdpClient.detach().catch(() => {});
    }

    if (novaPagina && !novaPagina.isClosed()) {
      await novaPagina.close().catch(() => {});
    }
  }
}

async function baixarAnexosElegiveisDaPaginaAtual(page, dataInicioPlanejado, pastaDestino, resultado) {
  const chavesVistas = new Set();
  const acoesPaginacaoVistas = new Set();
  const dataBaseComparacao = dataInicioPlanejado ? dataSomenteDia(dataInicioPlanejado) : null;

  async function processarLinhasPaginaAtual() {
    let linhas = page.locator('#tableFileUpload tr');
    if ((await linhas.count()) === 0) {
      linhas = page.locator("table[id*='tableFileUpload'] tr");
    }
    if ((await linhas.count()) === 0) {
      linhas = page.locator("table[id*='FileUpload'] tr");
    }
    if ((await linhas.count()) === 0) {
      linhas = page.locator('#tableFileUpload tbody tr');
    }
    if ((await linhas.count()) === 0) {
      linhas = page.locator("tr:has(a[id*='linkButtonDownloadAnexo']), tr:has(select[id*='dropDownListSituacaoAnexo']), tr:has(span[id*='textoTipoAnexo'])");
    }

    const totalLinhas = await linhas.count();

    for (let indice = 0; indice < totalLinhas; indice += 1) {
      const linha = linhas.nth(indice);
      const colunas = linha.locator('td, th');
      const totalColunas = await colunas.count();
      if (totalColunas === 0) {
        continue;
      }

      let descricao = totalColunas >= 1 ? textoLimpo(await colunas.nth(0).innerText()) : '';

      const formatoSpan = linha.locator("span[id*='textoTipoAnexo'], span[id*='TipoAnexo']");
      const formato = (await formatoSpan.count()) ? textoLimpo(await formatoSpan.first().innerText()) : '';

      const arquivoLink = linha.locator("a[id*='linkButtonDownloadAnexo'], a[id*='DownloadAnexo'], a[title*='Download']");
      if ((await arquivoLink.count()) === 0) {
        continue; // linha sem arquivo anexado, nada para baixar
      }

      const arquivoNome = textoLimpo(await arquivoLink.first().innerText()) || '(sem nome)';
      if (!descricao) {
        descricao = formato || arquivoNome;
      }

      const codanexo = (await arquivoLink.first().getAttribute('codanexo')) || '';
      const nomearquivoAttr = (await arquivoLink.first().getAttribute('nomearquivo')) || '';
      const downloadId = (await arquivoLink.first().getAttribute('id')) || '';

      const chaveAnexo = [codanexo, nomearquivoAttr, downloadId, descricao, arquivoNome].join('|');
      if (chavesVistas.has(chaveAnexo)) {
        continue;
      }
      chavesVistas.add(chaveAnexo);

      const historicoLink = linha.locator("a:has-text('(Histórico)'), a:has-text('Histórico'), a:has-text('Historico'), a[title*='Histórico'], a[title*='Historico'], a[id*='Historico']");

      let historico = [];
      try {
        historico = await extrairHistoricoArquivo(page, historicoLink);
      } catch {
        historico = [];
      }

      resultado.totalAnexosAnalisados += 1;

      let dataMaisRecente = null;
      let elegivel = false;
      for (const item of historico) {
        const dataItem = converterValorParaData(item.data);
        if (!dataItem) {
          continue;
        }
        if (!dataMaisRecente || dataItem > dataMaisRecente) {
          dataMaisRecente = dataItem;
        }
        if (dataBaseComparacao && dataSomenteDia(dataItem) >= dataBaseComparacao) {
          elegivel = true;
        }
      }

      if (!elegivel) {
        continue;
      }

      resultado.totalElegiveis += 1;

      try {
        const caminhoFinal = await baixarArquivoPorClique(
          page,
          arquivoLink,
          pastaDestino,
          arquivoNome || `${descricao || 'anexo'}.dat`
        );

        resultado.totalBaixados += 1;
        resultado.arquivosBaixados.push({
          descricao,
          arquivo: path.basename(caminhoFinal),
          data_postagem: formatarDataBR(dataMaisRecente),
        });
      } catch (erroDownload) {
        resultado.erros.push(
          `${descricao || arquivoNome}: ${erroDownload && erroDownload.message ? erroDownload.message : 'Falha ao baixar o arquivo.'}`
        );
      }
    }
  }

  async function clicarProximaPagina() {
    let assinaturaAntes = '';
    try {
      assinaturaAntes = textoLimpo(await page.locator('#tableFileUpload').first().innerText());
    } catch {
      try {
        assinaturaAntes = textoLimpo(await page.locator("table[id*='FileUpload']").first().innerText());
      } catch {
        assinaturaAntes = '';
      }
    }

    let linksPaginacao = page.locator("#tableFileUpload a[href*='Page$']");
    if ((await linksPaginacao.count()) === 0) {
      linksPaginacao = page.locator("table[id*='FileUpload'] a[href*='Page$']");
    }

    const totalLinks = await linksPaginacao.count();

    for (let idx = 0; idx < totalLinks; idx += 1) {
      const link = linksPaginacao.nth(idx);
      try {
        if (!(await link.isVisible())) {
          continue;
        }
      } catch {
        continue;
      }

      const href = ((await link.getAttribute('href')) || '').trim();
      const textoLink = textoLimpo(await link.innerText());
      const acao = href || textoLink;

      if (!acao || acoesPaginacaoVistas.has(acao)) {
        continue;
      }

      if (!acao.includes('Page$') && !['>', '»', 'Próx', 'Prox'].some((ch) => textoLink.includes(ch))) {
        continue;
      }

      acoesPaginacaoVistas.add(acao);

      try {
        await link.click({ force: true });
        await page.waitForLoadState('networkidle');
        await page.waitForTimeout(700);
      } catch {
        continue;
      }

      let assinaturaDepois = '';
      try {
        assinaturaDepois = textoLimpo(await page.locator('#tableFileUpload').first().innerText());
      } catch {
        try {
          assinaturaDepois = textoLimpo(await page.locator("table[id*='FileUpload']").first().innerText());
        } catch {
          assinaturaDepois = '';
        }
      }

      if (assinaturaDepois && assinaturaDepois !== assinaturaAntes) {
        return true;
      }
    }

    const seletoresProximo = [
      "#tableFileUpload a[href*='Page$Next']",
      "table[id*='FileUpload'] a[href*='Page$Next']",
      "#tableFileUpload a[title*='Próx']",
      "#tableFileUpload a[title*='Prox']",
      "#tableFileUpload a:has-text('Próximo')",
      "#tableFileUpload a:has-text('Proximo')",
      "table[id*='FileUpload'] a[title*='Próx']",
      "table[id*='FileUpload'] a[title*='Prox']",
      "table[id*='FileUpload'] a:has-text('Próximo')",
      "table[id*='FileUpload'] a:has-text('Proximo')",
      "#tableFileUpload a:has-text('>')",
      "#tableFileUpload a:has-text('»')",
      "table[id*='FileUpload'] a:has-text('>')",
      "table[id*='FileUpload'] a:has-text('»')",
    ];

    for (const seletor of seletoresProximo) {
      const link = page.locator(seletor);
      if ((await link.count()) && (await link.first().isVisible())) {
        try {
          const href = ((await link.first().getAttribute('href')) || '').trim();
          const textoLink = textoLimpo(await link.first().innerText());
          const acao = href || `fallback::${textoLink}`;
          if (acoesPaginacaoVistas.has(acao)) {
            continue;
          }
          acoesPaginacaoVistas.add(acao);
        } catch {
          // ignora
        }

        try {
          await link.first().click({ force: true });
          await page.waitForLoadState('networkidle');
          await page.waitForTimeout(700);
        } catch {
          continue;
        }

        let assinaturaDepois = '';
        try {
          assinaturaDepois = textoLimpo(await page.locator('#tableFileUpload').first().innerText());
        } catch {
          try {
            assinaturaDepois = textoLimpo(await page.locator("table[id*='FileUpload']").first().innerText());
          } catch {
            assinaturaDepois = '';
          }
        }

        if (assinaturaDepois && assinaturaDepois !== assinaturaAntes) {
          return true;
        }
      }
    }

    return false;
  }

  const maxPaginas = 12;
  for (let paginaAtual = 1; paginaAtual <= maxPaginas; paginaAtual += 1) {
    await processarLinhasPaginaAtual();
    const teveProxima = await clicarProximaPagina();
    if (!teveProxima) {
      break;
    }
  }
}

// Faz login, abre a NS informada, e baixa todos os anexos postados na data de início planejado
// ou depois dela, salvando em uma subpasta de Downloads no padrão NS_XXXX_DOCUMENTACAO_DDMMAAAA_HHMMSS.
async function baixarDocumentosPosterioresPlanejamento(numeroNs, dataInicioPlanejado, usuarioApr, senhaApr, { mostrarNavegador = false } = {}) {
  const browser = await criarBrowser({ mostrarNavegador });

  const nomePasta = montarNomePastaDownload(numeroNs);
  const pastaDestino = path.join(os.homedir(), 'Downloads', nomePasta);

  const resultado = {
    numeroNs,
    pastaDestino,
    totalAnexosAnalisados: 0,
    totalElegiveis: 0,
    totalBaixados: 0,
    arquivosBaixados: [],
    erros: [],
  };

  let context;
  let page;

  try {
    fs.mkdirSync(pastaDestino, { recursive: true });

    context = await browser.newContext({ acceptDownloads: true, downloadsPath: pastaDestino });
    page = await context.newPage();

    // Alguns anexos podem disparar um confirm()/alert() do navegador antes de liberar o arquivo;
    // sem esse handler, o Playwright cancela o diálogo automaticamente e o download nunca começa.
    page.on('dialog', (dialog) => {
      dialog.accept().catch(() => {});
    });

    await abrirNotaNoApr(page, numeroNs, usuarioApr, senhaApr);

    await garantirSecaoAnexosVisivel(page);
    await aguardarAnexosRenderizados(page);

    await page.waitForSelector("#tableFileUpload, table[id*='FileUpload'], #tableFileUpload tbody tr, table[id*='FileUpload'] tbody tr, tr:has(a[id*='linkButtonDownloadAnexo']), tr:has(a[id*='DownloadAnexo']), tr:has(select[id*='dropDownListSituacaoAnexo']), tr:has(select[id*='SituacaoAnexo'])", { state: 'visible', timeout: 15000 }).catch(() => {});

    // O executável do Chrome usado aqui não está disparando o evento "download" do Playwright de
    // forma confiável (confirmado: o servidor responde com Content-Disposition: attachment, mas o
    // evento nunca chega). Como contorno, instruímos o Chrome via CDP a salvar automaticamente
    // qualquer download diretamente nesta pasta — depois só verificamos quando o arquivo aparece.
    await habilitarDownloadViaCdp(page, pastaDestino);

    await baixarAnexosElegiveisDaPaginaAtual(page, dataInicioPlanejado, pastaDestino, resultado);
  } finally {
    if (context) {
      await context.close().catch(() => {});
    }
    await browser.close();
  }

  return resultado;
}

function renderPage({
  usuarioApr = '',
  mostrarNavegador = false,
  modoConsulta = 'individual',
  numeroNs = '',
  numerosNsLote = '',
  dados = null,
  anexos = [],
  resultadoLote = [],
  resultadoPlanilha = [],
  resultadoDownload = null,
  erro = '',
  qtdAnexos = 0,
  qtdHistoricos = 0,
  numeroInstalacao = '',
}) {
  const dadosRows = dados
    ? Object.entries(dados)
      .map(([campo, valor]) => `<tr><th>${escapeHtml(campo)}</th><td>${escapeHtml(valor)}</td></tr>`)
      .join('')
    : '';

  const resultadoLoteRows = resultadoLote
    .map((item) => `<tr><td>${escapeHtml(item.ns)}</td><td>${escapeHtml(item.nome_responsavel_tecnico)}</td><td>${escapeHtml(item.status)}</td></tr>`)
    .join('');
  const resultadoLoteExportB64 = resultadoLote.length
    ? Buffer.from(JSON.stringify(resultadoLote), 'utf8').toString('base64url')
    : '';

  const resultadoPlanilhaRows = resultadoPlanilha
    .map((item) => `<tr><td>${escapeHtml(item.nota)}</td><td>${escapeHtml(item.inicio_planejado)}</td><td class="possui-anexo-${escapeHtml(String(item.possui_anexo || '').toLowerCase())}">${escapeHtml(item.possui_anexo)}</td><td>${escapeHtml(item.data_postagem_mais_recente)}</td><td>${escapeHtml(item.observacao)}</td></tr>`)
    .join('');
  const resultadoPlanilhaExportB64 = resultadoPlanilha.length
    ? Buffer.from(JSON.stringify(resultadoPlanilha), 'utf8').toString('base64url')
    : '';

  const anexosRows = anexos
    .map((anexo) => {
      const historicoHtml = Array.isArray(anexo.historico) && anexo.historico.length
        ? `<ul class="hist-list">${anexo.historico.map((item) => `<li>${escapeHtml(item.texto || `${item.data || ''} - ${item.usuario || ''}`)}</li>`).join('')}</ul>`
        : '<div class="hist-vazio">Sem histórico disponível para esta linha.</div>';

      return `<tr><td>${escapeHtml(anexo.descricao)}</td><td>${escapeHtml(anexo.arquivo)}</td><td class="hist-cell">${historicoHtml}</td></tr>`;
    })
    .join('');

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Consulta APR por NS</title>
  <link rel="stylesheet" href="/static/botaprweb.css" />
</head>
<body>
  <div id="loading-overlay" class="loading-overlay" aria-hidden="true">
    <div class="loading-box">
      <div class="loading-spinner" role="status" aria-label="Carregando"></div>
      <div class="loading-text">Carregando informações, aguarde...</div>
    </div>
  </div>

  <div class="page-container">
    <div class="page-card">
      <h1 id="titulo">Consulta APR por NS (Node)</h1>

      <form method="post" action="/" class="form-grid" id="consulta-form" enctype="multipart/form-data">
        <div class="campo">
          <label for="tipo_pesquisa">Tipo de pesquisa</label>
          <select id="tipo_pesquisa" name="modo_consulta">
            <option value="individual" ${modoConsulta !== 'lote' && modoConsulta !== 'planilha' && modoConsulta !== 'download' ? 'selected' : ''}>NS individual</option>
            <option value="lote" ${modoConsulta === 'lote' ? 'selected' : ''}>Consulta em lote (responsáveis)</option>
            <option value="planilha" ${modoConsulta === 'planilha' ? 'selected' : ''}>Consulta por planilha (verificar anexos)</option>
            <option value="download" ${modoConsulta === 'download' ? 'selected' : ''}>Baixar documentos após início planejado</option>
          </select>
        </div>

        <div class="campo">
          <label for="usuario_apr">Matrícula APR</label>
          <input id="usuario_apr" name="usuario_apr" type="text" maxlength="20" value="${escapeHtml(usuarioApr)}" placeholder="Digite sua matrícula" required />
        </div>

        <!-- <div class="campo" style="position:relative;">
        //   <label for="senha_apr">Senha APR</label>
        //   <input id="senha_apr" name="senha_apr" type="password" maxlength="80" value="" placeholder="Digite sua senha" required style="padding-right:32px;" />
        //   <button type="button" id="toggle-senha-apr" style="position:absolute; top:32px; right:8px; background:none; border:none; cursor:pointer; width:24px; height:24px;" aria-label="Mostrar senha">
        //     <span id="icon-senha-apr" style="font-size:18px;">👁️</span>
        //   </button>
        // </div>-->

        

<div class="campo" style="position:relative;">
  <label for="senha_apr">Senha APR</label>

  <input id="senha_apr" name="senha_apr" type="password"
    maxlength="80"
    placeholder="Digite sua senha"
    required
    style="padding-right:32px;" />

  <button type="button" id="toggle-senha-apr"
    style="position:absolute; top:32px; right:8px; background:none; border:none; cursor:pointer;">
    👁️
  </button>
</div>



<div class="campo checkbox-campo">
  <input type="checkbox" id="lembrar_senha" />
  <label for="lembrar_senha">Lembrar senha nesta sessão</label>
</div>


        <div class="campo" id="campo-ns-individual">
          <label for="numero_ns">Número da NS</label>
          <input id="numero_ns" name="numero_ns" type="text" maxlength="20" value="${escapeHtml(numeroNs)}" placeholder="Digite o número da NS" />
        </div>

        <div class="campo" id="campo-ns-lote">
          <label for="numeros_ns_lote">Consulta em lote (NS)</label>
          <textarea id="numeros_ns_lote" name="numeros_ns_lote" rows="4" placeholder="Digite várias NS separadas por vírgula, espaço ou quebra de linha">${escapeHtml(numerosNsLote)}</textarea>
        </div>

        <div class="campo" id="campo-ns-planilha">
          <label for="planilha_notas">Planilha de notas (.xlsx)</label>
          <input id="planilha_notas" name="planilha_notas" type="file" accept=".xlsx,.xls" />
          <small class="ajuda-campo">
            A planilha deve ter as colunas <strong>Nota</strong> (número da NS) e <strong>InícPlanej</strong>
            (data de início planejado, ex.: 22/07/2026). Para cada nota, o sistema entra no APRWEB, varre os
            anexos e verifica se algum arquivo foi postado na data do início planejado ou depois dela.
          </small>
        </div>

        <div class="campo" id="campo-ns-download">
          <label for="numero_ns_download">Número da NS</label>
          <input id="numero_ns_download" name="numero_ns_download" type="text" maxlength="20" placeholder="Digite o número da NS" />
        </div>

        <div class="campo" id="campo-data-download">
          <label for="data_inicio_planejado_download">Início planejado</label>
          <input id="data_inicio_planejado_download" name="data_inicio_planejado_download" type="date" />
          <small class="ajuda-campo">
            Os documentos anexados na NS com data de postagem igual ou posterior a esta data serão baixados
            automaticamente para uma pasta dentro de "Downloads", nomeada
            <code>NS_&lt;numero&gt;_DOCUMENTACAO_DDMMAAAA_HHMMSS</code>.
          </small>
        </div>

        <div class="campo checkbox-campo">
          <label for="mostrar_navegador">Exibir navegação do bot</label>
          <input id="mostrar_navegador" name="mostrar_navegador" type="checkbox" value="1" ${mostrarNavegador ? 'checked' : ''} />
        </div>

        <div class="acoes-form">
          <button type="submit" class="btn-primario" id="btn-pesquisar">Pesquisar</button>
        </div>
      </form>

      ${erro ? `<div class="erro">${escapeHtml(erro)}</div>` : ''}

      ${dados ? `<div class="resumo">Resumo: Anexos = ${qtdAnexos} | Históricos = ${qtdHistoricos}</div>
      <div class="secao">
        <h2>Dados gerais</h2>
        <table><tbody>${dadosRows}</tbody></table>
      </div>` : ''}

      ${resultadoLote.length ? `<div class="secao">
        <h2>Resultado da consulta em lote</h2>
        <div class="acoes-lote">
          <form method="post" action="/exportar-lote-xlsx" class="form-exportar-lote">
            <input type="hidden" name="resultado_lote_b64" value="${escapeHtml(resultadoLoteExportB64)}" />
            <button type="submit" class="btn-primario" id="btn-exportar-lote">Exportar para Excel</button>
          </form>
        </div>
        <table id="tabela-resultado-lote">
          <thead><tr><th>NS</th><th>NOME RESPONSAVEL TECNICO</th><th>Status</th></tr></thead>
          <tbody>${resultadoLoteRows}</tbody>
        </table>
      </div>` : ''}

      ${resultadoPlanilha.length ? `<div class="secao">
        <h2>Resultado da verificação de anexos por planilha</h2>
        <div class="acoes-lote">
          <form method="post" action="/exportar-planilha-anexos-xlsx" class="form-exportar-lote">
            <input type="hidden" name="resultado_planilha_b64" value="${escapeHtml(resultadoPlanilhaExportB64)}" />
            <button type="submit" class="btn-primario" id="btn-exportar-planilha">Exportar para Excel</button>
          </form>
        </div>
        <table id="tabela-resultado-planilha">
          <thead><tr><th>Nota</th><th>InícPlanej</th><th>Possui_Anexo</th><th>Data postagem mais recente</th><th>Observação</th></tr></thead>
          <tbody>${resultadoPlanilhaRows}</tbody>
        </table>
      </div>` : ''}

      ${resultadoDownload ? `<div class="secao">
        <h2>Download de documentos posteriores ao início planejado</h2>
        <div class="resumo">
          NS ${escapeHtml(resultadoDownload.numeroNs)} — Pasta: <code>${escapeHtml(resultadoDownload.pastaDestino)}</code><br/>
          Anexos analisados: ${resultadoDownload.totalAnexosAnalisados} |
          Elegíveis (postados no início planejado ou depois): ${resultadoDownload.totalElegiveis} |
          Baixados com sucesso: ${resultadoDownload.totalBaixados}
        </div>
        ${resultadoDownload.arquivosBaixados.length ? `<table class="tabela-anexos">
          <thead><tr><th>Descrição</th><th>Arquivo salvo</th><th>Data de postagem</th></tr></thead>
          <tbody>${resultadoDownload.arquivosBaixados.map((item) => `<tr><td>${escapeHtml(item.descricao)}</td><td>${escapeHtml(item.arquivo)}</td><td>${escapeHtml(item.data_postagem)}</td></tr>`).join('')}</tbody>
        </table>` : '<div class="vazio">Nenhum arquivo elegível foi encontrado para download nesta NS.</div>'}
        ${resultadoDownload.erros.length ? `<div class="erro">Falha ao baixar alguns arquivos:<ul>${resultadoDownload.erros.map((msg) => `<li>${escapeHtml(msg)}</li>`).join('')}</ul></div>` : ''}
      </div>` : ''}

      <div class="versao">Versão: ${APP_VERSAO}</div>

      ${modoConsulta !== 'lote' && modoConsulta !== 'planilha' && modoConsulta !== 'download' ? `<div class="secao">
        <h2>Anexos e histórico de postagens</h2>
        ${anexos.length ? `<table class="tabela-anexos"><thead><tr><th>Descrição</th><th>Arquivo</th><th>Histórico de postagens</th></tr></thead><tbody>${anexosRows}</tbody></table>` : '<div class="vazio">Nenhum anexo/histórico retornado para esta NS.</div>'}
      </div>` : ''}
    </div>
  </div>

  <script>
(function () {

  const form = document.getElementById('consulta-form');
  const overlay = document.getElementById('loading-overlay');
  const button = document.getElementById('btn-pesquisar');

  const tipoPesquisa = document.getElementById('tipo_pesquisa');
  const campoUsuario = document.getElementById('usuario_apr');
  const campoSenha = document.getElementById('senha_apr');
  const lembrarSenha = document.getElementById('lembrar_senha');

  const campoNs = document.getElementById('numero_ns');
  const campoNsLote = document.getElementById('numeros_ns_lote');
  const blocoNsIndividual = document.getElementById('campo-ns-individual');
  const blocoNsLote = document.getElementById('campo-ns-lote');
  const blocoNsPlanilha = document.getElementById('campo-ns-planilha');
  const campoPlanilha = document.getElementById('planilha_notas');

  const blocoNsDownload = document.getElementById('campo-ns-download');
  const blocoDataDownload = document.getElementById('campo-data-download');
  const campoNsDownload = document.getElementById('numero_ns_download');
  const campoDataDownload = document.getElementById('data_inicio_planejado_download');

  const toggleSenhaBtn = document.getElementById('toggle-senha-apr');

  const modoConsultaAtual = ${JSON.stringify(modoConsulta)};

  if (!form) return;

  // =====================================
  // MOSTRAR / OCULTAR SENHA
  // =====================================

  if (toggleSenhaBtn) {
    toggleSenhaBtn.addEventListener('click', function () {

      if (campoSenha.type === 'password') {
        campoSenha.type = 'text';
        toggleSenhaBtn.textContent = '🙈';
      } else {
        campoSenha.type = 'password';
        toggleSenhaBtn.textContent = '👁️';
      }

    });
  }

  // =====================================
  // CARREGA DADOS SALVOS
  // =====================================

  try {

    const usuarioSalvo =
      localStorage.getItem('aprweb.usuario');

    const senhaSalva =
      localStorage.getItem('aprweb.senha');

    if (usuarioSalvo) {
      campoUsuario.value = usuarioSalvo;
    }

    if (senhaSalva) {
      campoSenha.value = senhaSalva;
      lembrarSenha.checked = true;
    }

  } catch (e) {
    console.error(e);
  }

  // =====================================
  // TROCA ENTRE INDIVIDUAL E LOTE
  // =====================================

  function atualizarModoPesquisa() {

    const modo =
      tipoPesquisa.value;

    blocoNsIndividual.style.display =
      modo === 'individual' ? '' : 'none';

    blocoNsLote.style.display =
      modo === 'lote' ? '' : 'none';

    blocoNsPlanilha.style.display =
      modo === 'planilha' ? '' : 'none';

    blocoNsDownload.style.display =
      modo === 'download' ? '' : 'none';

    blocoDataDownload.style.display =
      modo === 'download' ? '' : 'none';

    campoNs.required = modo === 'individual';
    campoNsLote.required = modo === 'lote';
    if (campoPlanilha) {
      campoPlanilha.required = modo === 'planilha';
    }
    if (campoNsDownload) {
      campoNsDownload.required = modo === 'download';
    }
    if (campoDataDownload) {
      campoDataDownload.required = modo === 'download';
    }

    button.textContent =
      modo === 'lote'
        ? 'Consultar lote'
        : modo === 'planilha'
          ? 'Verificar anexos da planilha'
          : modo === 'download'
            ? 'Baixar documentos'
            : 'Pesquisar NS';
  }

  tipoPesquisa.addEventListener(
    'change',
    atualizarModoPesquisa
  );

  tipoPesquisa.value =
    modoConsultaAtual === 'lote'
      ? 'lote'
      : modoConsultaAtual === 'planilha'
        ? 'planilha'
        : modoConsultaAtual === 'download'
          ? 'download'
          : 'individual';

  atualizarModoPesquisa();

  // =====================================
  // SALVA SENHA AUTOMATICAMENTE
  // =====================================

  campoSenha.addEventListener('input', function () {

    if (!lembrarSenha.checked)
      return;

    localStorage.setItem(
      'aprweb.senha',
      campoSenha.value
    );

  });

  // =====================================
  // CHECKBOX LEMBRAR SENHA
  // =====================================

  lembrarSenha.addEventListener('change', function () {

    if (lembrarSenha.checked) {

      localStorage.setItem(
        'aprweb.senha',
        campoSenha.value
      );

    } else {

      localStorage.removeItem(
        'aprweb.senha'
      );

    }

  });

  // =====================================
  // SUBMIT DO FORMULÁRIO
  // =====================================

  form.addEventListener('submit', function () {

    try {

      localStorage.setItem(
        'aprweb.usuario',
        campoUsuario.value
      );

      if (lembrarSenha.checked) {

        localStorage.setItem(
          'aprweb.senha',
          campoSenha.value
        );

      } else {

        localStorage.removeItem(
          'aprweb.senha'
        );

      }

    } catch (e) {
      console.error(e);
    }

    overlay.classList.add('is-visible');
    overlay.setAttribute('aria-hidden', 'false');

    button.disabled = true;
    button.textContent = 'Carregando...';

  });

})();
</script>




</body>
</html>`;
}

app.post('/exportar-lote-xlsx', (req, res) => {
  const payloadB64 = textoLimpo(req.body.resultado_lote_b64 || '');

  if (!payloadB64) {
    return res.status(400).send('Nenhum resultado de lote disponível para exportação.');
  }

  let resultadoLote = [];
  try {
    const json = Buffer.from(payloadB64, 'base64url').toString('utf8');
    const payload = JSON.parse(json);
    if (!Array.isArray(payload)) {
      return res.status(400).send('Payload de exportação inválido.');
    }
    resultadoLote = payload;
  } catch {
    return res.status(400).send('Falha ao decodificar os dados para exportação.');
  }

  const linhas = resultadoLote.map((item) => ({
    NS: textoLimpo(item.ns || ''),
    NOME_RESPONSAVEL_TECNICO: textoLimpo(item.nome_responsavel_tecnico || ''),
    STATUS: textoLimpo(item.status || ''),
  }));

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(linhas);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'ConsultaLote');

  const agora = new Date();
  const timestamp = `${agora.getFullYear()}${String(agora.getMonth() + 1).padStart(2, '0')}${String(agora.getDate()).padStart(2, '0')}-${String(agora.getHours()).padStart(2, '0')}${String(agora.getMinutes()).padStart(2, '0')}${String(agora.getSeconds()).padStart(2, '0')}`;
  const nomeArquivo = `resultado_consulta_lote_${timestamp}.xlsx`;
  const arquivoBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
  return res.send(arquivoBuffer);
});

app.post('/exportar-planilha-anexos-xlsx', (req, res) => {
  const payloadB64 = textoLimpo(req.body.resultado_planilha_b64 || '');

  if (!payloadB64) {
    return res.status(400).send('Nenhum resultado de verificação por planilha disponível para exportação.');
  }

  let resultadoPlanilha = [];
  try {
    const json = Buffer.from(payloadB64, 'base64url').toString('utf8');
    const payload = JSON.parse(json);
    if (!Array.isArray(payload)) {
      return res.status(400).send('Payload de exportação inválido.');
    }
    resultadoPlanilha = payload;
  } catch {
    return res.status(400).send('Falha ao decodificar os dados para exportação.');
  }

  const linhas = resultadoPlanilha.map((item) => ({
    Nota: textoLimpo(item.nota || ''),
    InicPlanej: textoLimpo(item.inicio_planejado || ''),
    Possui_Anexo: textoLimpo(item.possui_anexo || ''),
    Data_Postagem_Mais_Recente: textoLimpo(item.data_postagem_mais_recente || ''),
    Observacao: textoLimpo(item.observacao || ''),
  }));

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(linhas);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'VerificacaoAnexos');

  const agora = new Date();
  const timestamp = `${agora.getFullYear()}${String(agora.getMonth() + 1).padStart(2, '0')}${String(agora.getDate()).padStart(2, '0')}-${String(agora.getHours()).padStart(2, '0')}${String(agora.getMinutes()).padStart(2, '0')}${String(agora.getSeconds()).padStart(2, '0')}`;
  const nomeArquivo = `verificacao_anexos_planilha_${timestamp}.xlsx`;
  const arquivoBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
  return res.send(arquivoBuffer);
});

app.get('/', (req, res) => {
  const resultadoKey = textoLimpo(req.query.resultado || '');
  const numeroNsConsulta = textoLimpo(req.query.ns || '');

  let modoConsulta = 'individual';
  let numeroNs = '';
  let numerosNsLote = '';
  let erro = '';
  let dados = null;
  let anexos = [];
  let resultadoLote = [];
  let resultadoPlanilha = [];
  let resultadoDownload = null;
  let qtdAnexos = 0;
  let qtdHistoricos = 0;
  let numeroInstalacao = '';

  if (resultadoKey && RESULTADOS_PAGINA_CACHE.has(resultadoKey)) {
    const estado = RESULTADOS_PAGINA_CACHE.get(resultadoKey);
    RESULTADOS_PAGINA_CACHE.delete(resultadoKey);

    modoConsulta = estado.modoConsulta || 'individual';
    numeroNs = '';
    numerosNsLote = '';
    erro = estado.erro || '';
    dados = estado.dados || null;
    anexos = Array.isArray(estado.anexos) ? estado.anexos : [];
    resultadoLote = Array.isArray(estado.resultadoLote) ? estado.resultadoLote : [];
    resultadoPlanilha = Array.isArray(estado.resultadoPlanilha) ? estado.resultadoPlanilha : [];
    resultadoDownload = estado.resultadoDownload || null;
    qtdAnexos = Number(estado.qtdAnexos || 0);
    qtdHistoricos = Number(estado.qtdHistoricos || 0);
    numeroInstalacao = estado.numeroInstalacao || '';
  } else if (numeroNsConsulta && RESULTADOS_CACHE.has(numeroNsConsulta)) {
    const cache = RESULTADOS_CACHE.get(numeroNsConsulta);
    dados = cache.dados;
    anexos = cache.anexos;
    RESULTADOS_CACHE.delete(numeroNsConsulta);

    modoConsulta = 'individual';
    numeroNs = '';
    numerosNsLote = '';
  }

  if (dados && !numeroInstalacao) {
    qtdAnexos = anexos.length;
    qtdHistoricos = anexos.reduce((acc, item) => acc + ((item.historico || []).length), 0);
    numeroInstalacao = dados['NUMERO DE INSTALAÇÃO'] || dados['NUMERO DE INSTALACAO'] || '';
  }

  res.send(renderPage({
    modoConsulta,
    numeroNs,
    numerosNsLote,
    dados,
    anexos,
    resultadoLote,
    resultadoPlanilha,
    resultadoDownload,
    erro,
    qtdAnexos,
    qtdHistoricos,
    numeroInstalacao,
  }));
});

app.post('/', upload.single('planilha_notas'), async (req, res) => {
  const usuarioApr = textoLimpo(req.body.usuario_apr || '');
  const senhaApr = textoLimpo(req.body.senha_apr || '');
  const mostrarNavegador = req.body.mostrar_navegador === '1';
  const numeroNs = textoLimpo(req.body.numero_ns || '');
  const numerosNsLote = String(req.body.numeros_ns_lote || '').trim();
  const modoConsultaRaw = req.body.modo_consulta;
  const modoConsulta = modoConsultaRaw === 'lote'
    ? 'lote'
    : modoConsultaRaw === 'planilha'
      ? 'planilha'
      : modoConsultaRaw === 'download'
        ? 'download'
        : 'individual';

  let dados = null;
  let anexos = [];
  let resultadoLote = [];
  let resultadoPlanilha = [];
  let resultadoDownload = null;
  let erro = '';
  let qtdAnexos = 0;
  let qtdHistoricos = 0;
  let numeroInstalacao = '';

  if (!usuarioApr || !senhaApr) {
    erro = 'Informe matrícula e senha APR antes de pesquisar.';
  } else if (modoConsulta === 'lote') {
    const listaNs = parsearListaNs(numerosNsLote);
    if (!listaNs.length) {
      erro = 'Informe uma ou mais NS no campo de consulta em lote.';
    } else {
      for (const nsLote of listaNs) {
        try {
          const resposta = await consultarNsApr(nsLote, usuarioApr, senhaApr, { mostrarNavegador, incluirAnexos: false });
          const nomeResponsavel = normalizarNomeApenasTexto(resposta.dados['NOME RESPONSAVEL TECNICO'] || '');

          resultadoLote.push({
            ns: nsLote,
            nome_responsavel_tecnico: nomeResponsavel || '(não encontrado)',
            status: 'OK',
          });
        } catch (err) {
          resultadoLote.push({
            ns: nsLote,
            nome_responsavel_tecnico: '',
            status: err && err.message ? err.message : 'Erro na consulta',
          });
        }
      }
    }
  } else if (modoConsulta === 'planilha') {
    if (!req.file || !req.file.buffer || !req.file.buffer.length) {
      erro = 'Selecione um arquivo de planilha (.xlsx) com as colunas Nota e InícPlanej.';
    } else {
      let linhasPlanilha = [];
      try {
        const leitura = lerLinhasPlanilhaNotas(req.file.buffer);
        if (leitura.erro) {
          erro = leitura.erro;
        }
        linhasPlanilha = leitura.linhas;
      } catch {
        erro = 'Não foi possível ler a planilha enviada. Verifique se é um arquivo .xlsx válido.';
      }

      if (!erro && !linhasPlanilha.length) {
        erro = 'Nenhuma linha válida encontrada na planilha (colunas esperadas: Nota e InícPlanej).';
      }

      if (!erro) {
        for (const linha of linhasPlanilha) {
          const dataInicioPlanejado = converterValorParaData(linha.valorDataPlanejada);
          const inicioPlanejFormatado = formatarDataBR(dataInicioPlanejado) || textoLimpo(linha.valorDataPlanejada);

          try {
            // eslint-disable-next-line no-await-in-loop
            const resposta = await consultarNsApr(linha.nota, usuarioApr, senhaApr, { mostrarNavegador, incluirAnexos: true });
            const { possuiPosterior, dataMaisRecente } = verificarAnexoPosterior(resposta.anexos, dataInicioPlanejado);

            resultadoPlanilha.push({
              nota: linha.nota,
              inicio_planejado: inicioPlanejFormatado,
              possui_anexo: dataInicioPlanejado ? (possuiPosterior ? 'Sim' : 'Não') : 'Não',
              data_postagem_mais_recente: formatarDataBR(dataMaisRecente),
              observacao: dataInicioPlanejado
                ? ''
                : 'Data de InícPlanej não reconhecida na planilha para esta linha.',
            });
          } catch (errConsulta) {
            if (errConsulta instanceof NsSemResultadoError) {
              resultadoPlanilha.push({
                nota: linha.nota,
                inicio_planejado: inicioPlanejFormatado,
                possui_anexo: 'Não',
                data_postagem_mais_recente: '',
                observacao: 'Nota não localizada/sem retorno no APR — tratada como ausência de anexo.',
              });
            } else {
              resultadoPlanilha.push({
                nota: linha.nota,
                inicio_planejado: inicioPlanejFormatado,
                possui_anexo: 'Erro',
                data_postagem_mais_recente: '',
                observacao: errConsulta && errConsulta.message ? errConsulta.message : 'Erro na consulta.',
              });
            }
          }
        }
      }
    }
  } else if (modoConsulta === 'download') {
    const numeroNsDownload = textoLimpo(req.body.numero_ns_download || '');
    const dataInicioPlanejadoDownloadTexto = textoLimpo(req.body.data_inicio_planejado_download || '');

    if (!numeroNsDownload) {
      erro = 'Informe o número da NS para baixar os documentos.';
    } else if (!dataInicioPlanejadoDownloadTexto) {
      erro = 'Informe a data de início planejado para filtrar os documentos.';
    } else {
      const dataInicioPlanejadoDownload = converterValorParaData(dataInicioPlanejadoDownloadTexto);
      if (!dataInicioPlanejadoDownload) {
        erro = 'Não foi possível reconhecer a data de início planejado informada.';
      } else {
        try {
          resultadoDownload = await baixarDocumentosPosterioresPlanejamento(
            numeroNsDownload,
            dataInicioPlanejadoDownload,
            usuarioApr,
            senhaApr,
            { mostrarNavegador }
          );
        } catch (errDownload) {
          erro = errDownload && errDownload.message
            ? `Erro ao baixar documentos da NS: ${errDownload.message}`
            : 'Erro ao baixar documentos da NS.';
        }
      }
    }
  } else if (!numeroNs) {
    erro = 'Informe o número da NS para pesquisar.';
  } else {
    try {
      const resposta = await consultarNsApr(numeroNs, usuarioApr, senhaApr, { mostrarNavegador, incluirAnexos: true });
      dados = resposta.dados;
      anexos = resposta.anexos;
    } catch (err) {
      erro = err && err.message ? `Erro ao consultar NS no APR: ${err.message}` : 'Erro ao consultar NS no APR.';
    }
  }

  if (dados) {
    qtdAnexos = anexos.length;
    qtdHistoricos = anexos.reduce((acc, item) => acc + ((item.historico || []).length), 0);
    numeroInstalacao = dados['NUMERO DE INSTALAÇÃO'] || dados['NUMERO DE INSTALACAO'] || '';
  }

  const resultadoKey = `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
  RESULTADOS_PAGINA_CACHE.set(resultadoKey, {
    modoConsulta,
    dados,
    anexos,
    resultadoLote,
    resultadoPlanilha,
    resultadoDownload,
    erro,
    qtdAnexos,
    qtdHistoricos,
    numeroInstalacao,
  });

  return res.redirect(`/?resultado=${encodeURIComponent(resultadoKey)}`);
});

module.exports = app;

if (require.main === module) {
  app.listen(APP_PORTA, '127.0.0.1', () => {
    console.log(`Servidor Node em http://127.0.0.1:${APP_PORTA}`);
  });
}