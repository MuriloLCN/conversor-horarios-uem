document.getElementById('botao-conversao').addEventListener('click', () =>
    /*
    Função chamada quando o botão de conversão é clicado

    Recebe os arquivos do elemento seletor e chama a função de parse principal com ele
    */
    {
        let elementoInputArquivos = document.getElementById('file-input');

        if (elementoInputArquivos.files.length == 0)
        {
            alert("Nenhum arquivo enviado, por favor escolha o arquivo do seu PDF com o horário");
            return;
        }

        if (elementoInputArquivos.files.length >= 2)
        {
            alert("Mais de um arquivo enviado, por favor, escolha apenas um");
            return;
        }

        let arquivoEnviado = elementoInputArquivos.files[0];
        let reader = new FileReader();

        reader.onload = function(event)
        {
            let buffer = event.target.result;
            let listaTextosPaginas = converterPDFParaTexto(buffer);
        }

        reader.readAsArrayBuffer(arquivoEnviado);
    }
);

var inicio_s1 = undefined;
var fim_s1 = undefined;
var inicio_s2 = undefined;
var fim_s2 = undefined;

function converterPDFParaTexto(buffer)
{
    /*
        Obtém o texto cru do PDF com o horário página por página e passa para a função principal de parse
    */
    pdfjsLib.getDocument({data: buffer}).promise.then(async function(pdf) {
        let numeroPaginas = pdf.numPages;
        let textosPaginas = [];

        for (let i = 1; i <= numeroPaginas; i++)
        {
            textosPaginas[i - 1] = await processarPagina(i, pdf);
        }

        parsePrincipal(textosPaginas);

        return textosPaginas;
    });
}

async function processarPagina(numero, pdf)
{  
    /*
        Pega todo o texto de uma determinada página do PDF
    */
    let pagina = await pdf.getPage(numero);
    let texto = await pagina.getTextContent();
    
    let resultado = '';

    for (let item of texto.items)
    {
        resultado += item.str + ' ';
    }
    return resultado;
}

function parsePrincipal(textosPaginas)
{
    /*
        Função principal de parse
    */

    // Inicializando as matrizes de horários com elementos da classe Evento
    inicializar_matriz_horarios(0);
    inicializar_matriz_horarios(1);

    let contador = textosPaginas.length - 1;

    if (document.getElementById("optativa").checked)
    {
        contador -= 1;
    }

    for (let i = 0; i < contador; i++) // Não é necessária o parse da última página, por isso o -1
    {
        if (i == 0)
        {
            // Primeira página contém a tabela de de conteúdos e códigos, precisa de parse especial
            
            /*
            Nota: A primeira página parece sempre conter duas tabelas (uma para cada semestre)
            então não imagino que deva existir algum tipo de horário que contém apenas uma - ou se existem,
            ainda não conheco nenhum caso assim
            */
            parsearPrimeiraPagina(textosPaginas[i]);
        }
        else
        {
            // Demais páginas contém apenas as tabelas, parse depende apenas da quantidade de tabelas
            parseIntermediario(textosPaginas[i]);
        }
    }
    
    // Passa pelas matrizes de horários atualizando os dados com base nos códigos das matérias lidas
    atualizarMatrizesHorario();

    // Compactando matérias adjacentes para o ICS ficar bonitinho
    compactarHorarios();

    // Gerando texto do arquivo ICS
    let texto = gerarICS();

    // Download do arquivo gerado
    downloadFile(texto);
}

function parsearPrimeiraPagina(textoPrimeiraPagina)
{
    /*
        Realiza a leitura da primeira página do horário, que é diferente das demais por possuir um cabeçalho diferente
        e também conter a tabela de relação de matérias
    */

    // Remove asteriscos inúteis do texto
    textoPrimeiraPagina = makeAsterisksSingle(textoPrimeiraPagina);

    let partes = textoPrimeiraPagina.split("*");

    /*
    Partes:
    0: Cabeçalho com nome e talz
    1: Cabeçalho das tabelas
    2: Primeira linha das tabelas
    3: Segunda linha das tabelas
    4: Terceira linha das tabelas
    5: Quarta linha das tabelas
    6: [Possivel ou não] Quinta linha das tabelas -- aqui é importante checar se essa linha existe ou não,
    por causa que horários noturnos só tem quatro aulas e isso pode comprometer a ordem do parse

    Horários normais:
    7: Tabelinha com códigos e matérias
    8: Desnecessária
    9: Desnecessária

    Horários noturnos (?):    (?) => Ainda não confirmei
    6: Tabelinha com códigos e matérias
    7 e 8: Desnecessária
    */

    let primeira_linha = partes[2];
    let segunda_linha = partes[3];
    let terceira_linha = partes[4];
    let quarta_linha = partes[5];
    let quinta_linha = partes[6];

    let relacao_materias = partes[7];

    // Caso o horário seja de 4 matérias
    if (!quinta_linha.includes("|"))
    {
        relacao_materias = partes[6];
        quinta_linha = "";
    }

    montarTabelaDeMaterias(relacao_materias);

    /*
        Exemplo de linha:
        primeira_linha_s0 = "
    | 07:45|0| 6882-001|         | 6882-001|         |         |         |     | 13:30|0| 6889-001| 6897-001|         | 6897-001|         |         |
    M| 08:35|1|D67 -208 |         |D67 -208 |         |         |         |    T| 14:20|6|D67 -208 |D67 -208 |         |D67 -208 |         |         |
        "
    */
    
    // O zero aqui indica que é "adivinhado" que essa primeira página se refere ao primeiro semestre
    pegarDadosLinhaDaTabela(primeira_linha, 0);
    pegarDadosLinhaDaTabela(segunda_linha, 0);
    pegarDadosLinhaDaTabela(terceira_linha, 0);
    pegarDadosLinhaDaTabela(quarta_linha, 0);
    pegarDadosLinhaDaTabela(quinta_linha, 0);
}

function parseIntermediario(textoPaginaIntermediaria)
{
    /*
        Realiza o parse para páginas intermediárias do horário
    */
    textoPaginaIntermediaria = makeAsterisksSingle(textoPaginaIntermediaria);

    let partes = textoPaginaIntermediaria.split("*");
    
    if (partes.length == 0)
    {
        return;
    }    
    
    /*
    Partes

    0: Cabeçalho
    1: Cabeçalho das tabelas
    2: Primeira linha das colunas
    3: Segunda linha das colunas
    4: Terceira linha das colunas
    5: Quarta linha das colunas
    6: Quinta linha das colunas
    7: Desnecessário

    ou

    ...
    5: Quarta linha das colunas
    6: Desnecessário
    */
    
    let primeira_linha = partes[2];
    let segunda_linha = partes[3];
    let terceira_linha = partes[4];
    let quarta_linha = partes[5];
    let quinta_linha = partes[6];
    
    // O um aqui significa que é "adivinhado" que a tabela se refere ao segundo semestre, caso seja encontrada alguma
    // matéria anual
    pegarDadosLinhaDaTabela(primeira_linha, 1);
    pegarDadosLinhaDaTabela(segunda_linha, 1);
    pegarDadosLinhaDaTabela(terceira_linha, 1);
    pegarDadosLinhaDaTabela(quarta_linha, 1);
    if (quinta_linha.includes("|"))
    {
        pegarDadosLinhaDaTabela(quinta_linha, 1);
    }
}

function gerarMatrizDados()
{
    /*
        Monta a matriz de dados para gerar o arquivo excel
    */
    try {
        let dados = [];
        dados.push(["S1","Seg","Ter","Qua","Qui","Sex", "Sab"]);
        for (let i = 0; i < 15; i++)
        {
            let nova_linha = [];
            nova_linha.push(tabela_horarios_inicio[String(i)]);
            for (let j = 0; j < 6; j++)
            {
                if (matriz_de_horarios_s0[i][j].materia == undefined)
                {
                    nova_linha.push('');
                }
                else 
                {
                    nova_linha.push(matriz_de_horarios_s0[i][j].materia);
                }
            }
            dados.push(nova_linha);
        }
        dados.push(["", "", "", "", "", "", ""]);
        dados.push(["S2","Seg","Ter","Qua","Qui","Sex", "Sab"]);
        for (let i = 0; i < 15; i++)
        {
            let nova_linha = [];
            nova_linha.push(tabela_horarios_inicio[String(i)]);
            for (let j = 0; j < 6; j++)
            {
                if (matriz_de_horarios_s1[i][j].materia == undefined)
                {
                    nova_linha.push('');
                }
                else 
                {
                    nova_linha.push(matriz_de_horarios_s1[i][j].materia);
                }
            }
            dados.push(nova_linha);
        }

        return dados;
    }
    catch (e)
    {
        alert("Nenhum arquivo ainda foi gerado, certifique-se de fazer o upload do seu arquivo de horário e de gerar o ICS nos passos anteriores");
        return undefined;
    }
}

document.getElementById("generateExcel").addEventListener("click", function() {
    try {

        const dados = gerarMatrizDados();
        if (!dados) return;

        const estilos = {
            cabecalhoDias: {
                fill: { fgColor: { rgb: "FFFF00" } },
            },
            colunaHorarios: {
                fill: { fgColor: { rgb: "EEEEEE" } }
            },
            bordas: {
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            }
        }

        var ws = XLSX.utils.aoa_to_sheet(dados);
        
        ajustarLarguraColunas(ws, dados);
        
        aplicarEstilos(ws, dados, estilos);

        exportarParaExcel(ws, "horario.xlsx");
        
    } catch (error) {
        console.error("Erro ao gerar Excel:", error);
        alert("Ocorreu um erro ao gerar o arquivo Excel.");
    }
});

function ajustarLarguraColunas(worksheet, dados) {
    let objectMaxLen = [];

    for (let col = 0; col < dados[0].length; col++) {
        objectMaxLen[col] = 0;
    }
    for (let row = 0; row < dados.length; row++) {
        for (let col = 0; col < dados[row].length; col++) {
            const cellValue = dados[row][col] ? dados[row][col].toString().trim() : "";
            
            if (cellValue.length > objectMaxLen[col]) {
                objectMaxLen[col] = cellValue.length + 6 ;
            }
        }
    }

    const larguras = objectMaxLen.map(w => { return { width: w } });

    worksheet['!cols'] = larguras;
}

function aplicarEstilos(worksheet, dados, estilos) {
    
    for (let linha = 0; linha < dados.length; linha++) {
        const cell = worksheet[XLSX.utils.encode_cell({ r: linha, c: 0 })] || {};
        if (dados[linha][0]) {
            cell.s = { ...estilos.colunaHorarios, ...estilos.bordas };
        }
    }

    const linhasCabecalho = [0, dados.findIndex(row => row[0] === "S2")];

    linhasCabecalho.forEach(linha => {
        for (let col = 0; col < dados[linha].length; col++) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: linha, c: col })] || {};
            cell.s = { ...estilos.cabecalhoDias, ...estilos.bordas };
        }
    });

    linhasCabecalho.forEach(linhaCabecalho => {
        const inicio = linhaCabecalho + 1;
        const fim = Math.min(linhaCabecalho + 14, dados.length);
        
        for (let linha = inicio; linha <= fim; linha++) {
            for (let col = 0; col < dados[linha].length; col++) {
                const cell = worksheet[XLSX.utils.encode_cell({ r: linha, c: col })] || {};
                cell.s = { ...(cell.s || {}), ...estilos.bordas };
            }
        }
    });
}

function exportarParaExcel(worksheet, nomeArquivo) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, worksheet, "Horário");
    
    const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([buffer], { type: 'application/octet-stream' });
    
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = nomeArquivo;
    link.click();
}

class Evento
{
      /*
          Classe de dados para organizar os dados das matérias
      */
      /*
      dia -> 0: seg, 1: ter, 2: qua, 3: qui, 4: sex, 5: sab
      horario -> 0: 0745-0835, 1:0835-0925, ...
      semestre -> 0 = primeiro, 1 = segundo, 2 = anual
      */
      constructor(dia, horario, semestre)
      {
          this.codigo = undefined;
          this.materia = undefined;
          this.dia_da_semana = dia;
          this.periodo = horario;
          this.semestre = semestre;
          this.nulo = true;
          this.local = undefined;
          this.data_inicio = undefined;
          this.data_termino = undefined;
          this.horario_inicio = undefined;
          this.horario_termino = undefined;
      }
}

// Matrizes, onde cada elemento representa um horário e dia da semana e os elementos são da classe "Evento"
var matriz_de_horarios_s0 = [];
var matriz_de_horarios_s1 = [];

// Tabela que relaciona o código de cada matéria com os dados da mesma
// Exemplo: tabela_materias["6887"] -> Circuitos Digitais 1, data de inicio: DD/MM/YYYY, etc.
var tabela_materias = {};

// Tabelas para converter os índices das matrizes de horários para os seus horários de inicio e fim
var tabela_horarios_inicio = {
  "0":"7:45",
  "1":"8:35",
  "2":"9:40",
  "3":"10:30",
  "4":"11:20",
  "5":"13:30",
  "6":"14:20",
  "7":"15:20",
  "8":"16:10",
  "9":"17:00",
  "10": "19:30",
  "11": "20:20",
  "12": "21:20",
  "13": "22:10"
};

var tabela_horarios_termino = {
  "0":"8:35",
  "1":"9:20",
  "2":"10:30",
  "3":"11:20",
  "4":"12:10",
  "5":"14:20",
  "6":"15:10",
  "7":"16:10",
  "8":"17:00",
  "9":"17:50",
  "10": "20:20",
  "11": "21:10",
  "12": "22:10",
  "13": "23:00"
};

var tabela_horarios_inverso = {
  "07:45": 0,
  "08:35": 1,
  "09:40": 2,
  "10:30": 3,
  "11:20": 4,
  "13:30": 5,
  "14:20": 6,
  "15:20": 7,
  "16:10": 8,
  "17:00": 9,
  "19:30": 10,
  "20:20": 11,
  "21:20": 12,
  "22:10": 13
};

function atualizarMatrizesHorario()
{
    /*
        Passa por cada elemento das duas matrizes de horário trocando
        o código da matéria pelo seu nome e colocando as suas datas de início e fim
    */
    // Primeiro semestre
    for (let i = 0; i < 15; i++)
    {
        for (let p = 0; p < 6; p++)
        {
            // Dados obtidos a partir do código da matéria
            let dados = tabela_materias[matriz_de_horarios_s0[i][p].codigo]; 
            if (dados == undefined) 
            {
                continue;
            }

            let dt_inicio = dados["data_inicio"];
            let dt_fim = dados["data_fim"];

            if (dados["semestre"] == 2)
            {
                dt_inicio = inicio_s1;
                dt_fim = fim_s1;
            }

            matriz_de_horarios_s0[i][p].materia = dados["nome"];
            matriz_de_horarios_s0[i][p].data_inicio = dt_inicio;
            matriz_de_horarios_s0[i][p].data_termino = dt_fim;
        }
    }

    // Segundo semestre
    for (let j = 0; j < 15; j++)
    {
        for (let q = 0; q < 6; q++)
        {
        let dados = tabela_materias[matriz_de_horarios_s1[j][q].codigo]; 
        if (dados == undefined) 
        {
            continue;
        }


        let dt_inicio = dados["data_inicio"];
        let dt_fim = dados["data_fim"];

        if (dados["semestre"] == 2)
        {
            dt_inicio = inicio_s2;
            dt_fim = fim_s2;
        }

        matriz_de_horarios_s1[j][q].materia = dados["nome"];
        matriz_de_horarios_s1[j][q].data_inicio = dt_inicio;
        matriz_de_horarios_s1[j][q].data_termino = dt_fim;
        }
    }
}

function inicializar_matriz_horarios(semestre)
{
    /*
        Inicializa a matriz de horários com elementos da classe "Evento" novos
    */
    for (let i = 0; i < 15; i++)
    {
        let horario = [];
        for (let p = 0; p < 6; p++)
        {
            let evento = new Evento(p, i, semestre);
            horario.push(evento);
        }

        if (semestre == 0)
        {
            matriz_de_horarios_s0.push(horario);
        }
        if (semestre == 1)
        {
            matriz_de_horarios_s1.push(horario);
        }
    }
}

function verificaSemestreMateria(codigo)
{
    /*
        Verifica o semestre no qual uma matéria é ministrada com base no seu código
    */
    codigo = replaceAll(codigo, ' ', '');
    let dados = tabela_materias[codigo];

    if (dados === undefined)
    {
        return undefined;
    }
    return parseInt(dados["semestre"]);
}

function pegarDadosMaterias(texto_materia)
{
    /*
        Obtem os dados de cada matéria a partir do texto da linha no horário
        Exemplo:
        6882   1 CIRCUITOS DIGITAIS II               DIN S1 26/06/23 24/10/23  17
        Irá montar:
        dados = {
            "data_fim": "24/10/23",
            "data_inicio": "26/06/23",
            "nome": "CIRCUITOS DIGITAIS II",
            "codigo": "6882",
            "semestre": "0",
        }
    */

    let dados = {};

    let partes = replaceAll(texto_materia, '  ', ' ').split(' ');
    
    if (partes[0] == '')
    {
        partes.shift();
    }
    /*
        0: codigo
        1: turma
        3 - ?: Nome
        tamanho - 1: limite faltas
        tamanho - 2: data termino
        tamanho - 3: data inicio
        tamanho - 4: semestre
        tamanho - 5: departamento

        OU, caso seja um horário referente ao segundo semestre:

        0: codigo
        1: turma
        tamanho - 1: semestre
        tamanho - 2: departamento
        
        Caso seja segundo semestre, as matérias já cursadas aparecem sem as datas de início e término e também sem limite de faltas
        nesses casos não é necessária colocá-las no horário a ser gerado
    */
    
    let tamanho = partes.length;

    if (!partes[tamanho - 2].includes('/'))
    {
        return undefined;
    }


    dados["nome"] = "";

    for (let i = 0; i < tamanho; i++)
    {
        if (i == 0)
        {
            dados["codigo"] = partes[i];
        }
        else if (i == 1)
        {
            dados["turma"] = partes[i];
        }
        else if (i == tamanho - 1)
        {
            dados["limite_faltas"] = partes[i];
        }
        else if (i == tamanho - 2)
        {
            dados["data_fim"] = partes[i];
        }
        else if (i == tamanho - 3)
        {
            dados["data_inicio"] = partes[i];
        }
        else if (i == tamanho - 4)
        {
            if (partes[i].includes("1"))
            {
                dados["semestre"] = 0;
            }
            else if (partes[i].includes("2"))
            {
                dados["semestre"] = 1;
            }
            else
            {
                dados["semestre"] = 2; // Materias anuais
            }
        }
        else if (i == tamanho - 5)
        {
            dados["departamento"] = partes[i];
        }
        else
        {
            dados["nome"] = dados["nome"] + " " + partes[i];
        }
    }

    if (dados["semestre"] == 0)
    {
        // console.log("Matéria era do primeiro semestre");
        if (inicio_s1 === undefined)
        {
            inicio_s1 = dados["data_inicio"];
        }
        if (fim_s1 === undefined)
        {
            fim_s1 = dados["data_fim"];
        }
    }
    else if (dados["semestre"] == 1)
    {
        // console.log("Matéria era do segundo semestre");
        if (inicio_s2 === undefined)
        {
            inicio_s2 = dados["data_inicio"];
        }
        if (fim_s2 === undefined)
        {
            fim_s2 = dados["data_fim"];
        }
    }

    return dados;
}

function montarTabelaDeMaterias(texto_com_materias)
{
  /*
      Lê a seção do pdf com os códigos e os nome + dados das matérias e salva elas na tabela_materias
      Exemplo:
      6882   1 CIRCUITOS DIGITAIS II               DIN S1 26/06/23 24/10/23  17
      6883   1 LINGUAGENS FORMAIS E AUTOMATOS      DIN S1 26/06/23 24/10/23  25
      6885   2 PROC.DE SOFT.E ENG. DE REQUISITOS   DIN S1 26/06/23 24/10/23  17
      6887   1 ARQUIT.E ORGANIZ.DE COMPUTADORES I  DIN S2 01/11/23 16/03/24  25
      6889   1 PROJETO E ANALISE DE ALGORITMOS     DIN S1 26/06/23 24/10/23  25
      [...]

      Montaria:
      tabela_materias = {
          "6882": {"nome": "CIRCUITOS DIGITAIS II", "inicio": "...", ...},
          ...
      }
  */

    // Divide as partes por quebra de linha
    let partes = texto_com_materias.split("\n");
    let inicio;

    // Verificação experimental de onde as matérias efetivamente começam no vetor "partes"
    if (partes[0] == "")
    {
        inicio = 4;
    }
    else
    {
        inicio = 3;
    }

    let i = inicio;
    while (partes[i] !== " " && i <= partes.length)
    {
        let dados = pegarDadosMaterias(partes[i]);

        if (dados != undefined) 
        {
            tabela_materias[dados["codigo"]] = dados;
            
        }
        i += 1;
    }
}

function replaceAll(text, str, replr)
{
    /*
        Apenas troca todas as ocorrências de uma substring em um texto por outra substring
    */
    while (text.includes(str))
    {
        text = text.replace(str, replr);
    }
    return text;
}

function makeAsterisksSingle(pdf_text)
{
    /*
        Dá uma limpada nos asteriscos do pdf com os horários
    */
    pdf_text = replaceAll(pdf_text, "*-", "*");
    pdf_text = replaceAll(pdf_text, "* -", "*");
    pdf_text = replaceAll(pdf_text, "*-", "*");
    pdf_text = replaceAll(pdf_text, "* *", "*");
    pdf_text = replaceAll(pdf_text, "**", "*");
    pdf_text = replaceAll(pdf_text, "* ", "*");
    pdf_text = replaceAll(pdf_text, "*R", "*");
    pdf_text = replaceAll(pdf_text, "**", "*");
    pdf_text = replaceAll(pdf_text, "  ", " ");
    return pdf_text;
}

function pegarDadosLinhaDaTabela(texto_da_linha, chute_semestre)
{
    /*
        Lê uma linha de uma ou duas tabelas de horários e atualiza a matriz de horários
        correspondente

        Exemplo (no primeiro semestre):
        | 07:45|0| 6882-001|         | 6882-001|         |         |         |     | 13:30|0| 6889-001| 6897-001|         | 6897-001|         |         |
        M| 08:35|1|D67 -208 |         |D67 -208 |         |         |         |    T| 14:20|6|D67 -208 |D67 -208 |         |D67 -208 |         |         |
        Colocaria na matriz do semestre 0 os locais e códigos das matérias
    */

    if (texto_da_linha === '')
    {
        return;
    }

    
    // Dividindo a parte de cima e de baixo da linha
    let partes = texto_da_linha.split("\n");

    if (partes[0] == '')
    {
        partes.shift();
    }

    let materias, locais;
    materias = partes[0];
    locais = partes[1];

    // Separando cada parte pelos seus '|'
    colunas_materias = materias.split("|");
    colunas_locais = locais.split("|");

    if (texto_da_linha.split('|').length - 1 < 36)
    {
        // --------------------------
        // Casos com uma tabela apenas
        // --------------------------

        /*
            " | 19:30|1| | | | | 9909-031| |"
            " N| 20:20|3| | | | |D67 -108 | |"
        */

        /*
        Colunas matérias
        0: --
        1: Horário início
        2: --
        3: Matéria segunda
        4: Matéria terça
        5: Matéria quarta
        6: Matéria quinta
        7: Matéria sexta
        8: Matéria sábado
        9: --

        Colunas locais
        0: --
        1: Horário término
        2: --
        3: Local segunda
        4: Local terça
        5: Local quarta
        6: Local quinta
        7: Local sexta
        8: Local sábado
        */

        let horario = colunas_materias[1];
        horario = replaceAll(horario, ' ', '');
        horario = tabela_horarios_inverso[horario];

        for (let i = 3; i < 9; i++)
        {
            let codigo = colunas_materias[i];
            codigo = replaceAll(codigo, ' ', '').split('-')[0];

            let semestre = verificaSemestreMateria(codigo);

            // Matérias anuais -> chutar semestre do qual a tabela se trata
            if (semestre === 2)
            {
                semestre = chute_semestre;
            }
            
            if (semestre == 0)
            {
                matriz_de_horarios_s0[horario][i - 3].codigo = codigo;
                matriz_de_horarios_s0[horario][i - 3].local = colunas_locais[i];
                matriz_de_horarios_s0[horario][i - 3].horario_inicio = tabela_horarios_inicio[String(horario)];
                matriz_de_horarios_s0[horario][i - 3].horario_termino = tabela_horarios_termino[String(horario)];
            }
            else
            {
                matriz_de_horarios_s1[horario][i - 3].codigo = codigo;
                matriz_de_horarios_s1[horario][i - 3].local = colunas_locais[i];
                matriz_de_horarios_s1[horario][i - 3].horario_inicio = tabela_horarios_inicio[String(horario)];
                matriz_de_horarios_s1[horario][i - 3].horario_termino = tabela_horarios_termino[String(horario)];
            }
        }   
        return;
    }


    /*
    Colunas materias:
    0: " "
    1: Horario inicio tabela 1
    2: --
    3: Materia segunda
    4: Materia terca
    5: Materia quarta
    6: Materia quinta
    7: Materia sexta
    8: Materia sábado
    9: --
    10: Horario inicio tabela 2
    11: --
    12: Materia segunda
    13: Materia terca
    14: Materia quarta
    15: Materia quinta
    16: Materia sexta
    17: Materia sábado
    18: --

    Colunas locais:
    0: --
    1: Horario fim tabela 1
    2: --
    3: Local segunda
    4: Local terca 
    5: Local quarta
    6: Local quinta
    7: Local sexta
    8: Local sábado
    9: --
    10: Horario fim tabela 2
    11: --
    12: Local segunda
    13: Local terca 
    14: Local quarta
    15: Local quinta
    16: Local sexta
    17: Local sábado
    18: --
    */

    let horario_tabela_1 = colunas_materias[1];
    horario_tabela_1 = replaceAll(horario_tabela_1, ' ', '');
    horario_tabela_1 = tabela_horarios_inverso[horario_tabela_1];

    
    let horario_tabela_2 = colunas_materias[10];
    horario_tabela_2 = replaceAll(horario_tabela_2, ' ', '');
    horario_tabela_2 = tabela_horarios_inverso[horario_tabela_2];
    
    // Primeira tabela

    for (let i = 3; i < 9; i++)
    {
        let codigo = colunas_materias[i];
        // Converte " 1234-4 " em só "1234" que é o código relevante
        codigo = replaceAll(codigo, ' ', '').split('-')[0];

        if (codigo === '')
        {
            continue;
        }

        let semestre = verificaSemestreMateria(codigo);

        if (semestre === 2)
        {
            semestre = chute_semestre;
        }
        
        if (semestre == 0)
        {
            matriz_de_horarios_s0[horario_tabela_1][i - 3].codigo = codigo;
            matriz_de_horarios_s0[horario_tabela_1][i - 3].local = colunas_locais[i];
            matriz_de_horarios_s0[horario_tabela_1][i - 3].horario_inicio = tabela_horarios_inicio[String(horario_tabela_1)];
            matriz_de_horarios_s0[horario_tabela_1][i - 3].horario_termino = tabela_horarios_termino[String(horario_tabela_1)];
        }
        else
        {
            matriz_de_horarios_s1[horario_tabela_1][i - 3].codigo = codigo;
            matriz_de_horarios_s1[horario_tabela_1][i - 3].local = colunas_locais[i];
            matriz_de_horarios_s1[horario_tabela_1][i - 3].horario_inicio = tabela_horarios_inicio[String(horario_tabela_1)];
            matriz_de_horarios_s1[horario_tabela_1][i - 3].horario_termino = tabela_horarios_termino[String(horario_tabela_1)];
        }
    }   

    // Segunda tabela 

    for (let i = 12; i < 18; i++)
    {
        let codigo = colunas_materias[i];
        // Converte " 1234-4 " em só "1234" que é o código relevante
        codigo = replaceAll(codigo, ' ', '').split('-')[0];

        if (codigo === '')
        {
            continue;
        }

        let semestre = verificaSemestreMateria(codigo);

        if (semestre === 2)
        {
            semestre = chute_semestre;
        }

        if (semestre == 0)
        {
            matriz_de_horarios_s0[horario_tabela_2][i - 12].codigo = codigo;
            matriz_de_horarios_s0[horario_tabela_2][i - 12].local = colunas_locais[i];
            matriz_de_horarios_s0[horario_tabela_2][i - 12].horario_inicio = tabela_horarios_inicio[String(horario_tabela_2)];
            matriz_de_horarios_s0[horario_tabela_2][i - 12].horario_termino = tabela_horarios_termino[String(horario_tabela_2)];
        }
        else
        {
            matriz_de_horarios_s1[horario_tabela_2][i - 12].codigo = codigo;
            matriz_de_horarios_s1[horario_tabela_2][i - 12].local = colunas_locais[i];
            matriz_de_horarios_s1[horario_tabela_2][i - 12].horario_inicio = tabela_horarios_inicio[String(horario_tabela_2)];
            matriz_de_horarios_s1[horario_tabela_2][i - 12].horario_termino = tabela_horarios_termino[String(horario_tabela_2)];
        }
    }
}

function converterHorarios(horario_normal) 
{
    // Converte o formato de horários para o arquivo ics
    //7:45 -> 074500      12:00 -> 120000
    let partes = horario_normal.split(":");
    let texto = "";
    if (Number(partes[0]) <= 9)
    {
        texto = texto + "0";
    }
    texto = texto + partes[0];
    texto = texto + partes[1];
    texto = texto + "00";
    return texto;
}

function primeiraAula(evento)
{
    // Calcula quando a primeira aula de uma matéria vai ocorrer
    // Por exemplo, se o semestre começou dia 1 de um mês e esse dia caiu em uma segunda,
    // a primeira aula de uma matéria que fica na quarta-feira seria no dia 3, assim usamos
    // essa nova data pra impedir que no primeiro dia do calendário tenha uma cópia de todas
    // as matérias possíveis, e apenas as que realmente caem naquele dia
    if (evento.materia == undefined)
    {
        return undefined;
    }
    let data_original = evento.data_inicio;

    let dia_da_semana = evento.dia_da_semana;

    let partes = data_original.split("/");

    let dia = partes[0];
    let mes = partes[1];
    let ano = "20" + partes[2];

    let data_inicio = new Date(Number(ano),Number(mes) - 1,Number(dia), 0, 0, 0, 0);

    while (data_inicio.getDay() - 1 !== dia_da_semana)
    {
        data_inicio.setDate(data_inicio.getDate() + 1);
    }

    return data_inicio.toLocaleDateString();
}

function converterDatas(data_normal, adicionar_20)
{
    // Converte datas para o formato do arquivo ICS
    // 16/03/24 -> 20240316

    let partes = data_normal.split("/");
    partes = partes.reverse();
    let texto = partes.join("");
    if (adicionar_20 == true)
    {
        texto = "20" + texto;
    }
    return texto;
}

// Horario com os eventos de maneira compactada e desordenada (a ordem não importa na hora de importar os eventos)
var horario_compactado = [];

function compactarHorarios()
{
  /*
      Lê as matrizes de horários e compacta elas 
  */
  // Para cada dia da semana
  for (let j = 0; j < 6; j++)
  {
      let evento_atual = matriz_de_horarios_s0[0][j];
      // Para cada horário 
      for (let i = 1; i < 14; i++)
      {
          // Essa monstrosidade vê se deve-se ou não compactar os horários e/ou colocar eles no horário compactado
          if (matriz_de_horarios_s0[i][j] == undefined)
          {
              if (evento_atual == undefined || evento_atual.materia == undefined)
              {
                  continue;
              }
              horario_compactado.push(evento_atual);
              evento_atual = undefined;
              continue;
          }

          if (evento_atual == undefined || evento_atual.materia == undefined)
          {
              evento_atual = matriz_de_horarios_s0[i][j];
              continue;
          }

          if (matriz_de_horarios_s0[i][j].materia == evento_atual.materia)
          {
              evento_atual.horario_termino = matriz_de_horarios_s0[i][j].horario_termino;
          }
          else
          {
              horario_compactado.push(evento_atual);
              evento_atual = matriz_de_horarios_s0[i][j];
          }
      }

      if (evento_atual == undefined || evento_atual.materia == undefined)
      {
          continue;
      }
      else
      {
          horario_compactado.push(evento_atual);
      }
  }
  
  for (let j = 0; j < 6; j++)
  {
      let evento_atual = matriz_de_horarios_s1[0][j];
      for (let i = 1; i < 14; i++)
      {
          if (matriz_de_horarios_s1[i][j] == undefined)
          {
              if (evento_atual == undefined)
              {
                  continue;
              }
              horario_compactado.push(evento_atual);
              evento_atual = undefined;
              continue;
          }

          if (evento_atual == undefined || evento_atual.materia == undefined)
          {
              evento_atual = matriz_de_horarios_s1[i][j];
              continue;
          }

          if (matriz_de_horarios_s1[i][j].materia == evento_atual.materia)
          {
              evento_atual.horario_termino = matriz_de_horarios_s1[i][j].horario_termino;
          }
          else
          {
              horario_compactado.push(evento_atual);
              evento_atual = matriz_de_horarios_s1[i][j];
          }
      }

      if (evento_atual == undefined || evento_atual.materia == undefined)
      {
          continue;
      }
      else
      {
          horario_compactado.push(evento_atual);
      }
  }

  //console.log(horario_compactado);
}

// Para converter índices em dias da semana para o formato ICS
var diaICS = {
  "0": "MO",
  "1": "TU",
  "2": "WE",
  "3": "TH",
  "4": "FR",
  "5": "SA"
};

function gerarTextoICSEvento(evento)
{
  // Gera o texto ICS para um evento, exemplo:

  // BEGIN:VEVENT
  // DTSTART;TZID=America/Sao_Paulo:20231116T160000
  // RRULE:FREQ=WEEKLY;BYDAY=TH;UNTIL=20231129T190000Z
  // DTEND;TZID=America/Sao_Paulo:20231116T190000
  // SUMMARY:TesteNome2
  // DESCRIPTION:TESTEEE
  // LOCATION:DINUEM2
  // END:VEVENT

  if (evento.materia == undefined)
  {
      return "";
  }

  let texto = "BEGIN:VEVENT\n";

  texto = texto + "DTSTART;TZID=America/Sao_Paulo:" + converterDatas(primeiraAula(evento)) + "T" + converterHorarios(evento.horario_inicio) + "\n";
  texto = texto + "RRULE:FREQ=WEEKLY;BYDAY="+diaICS[String(evento.dia_da_semana)]+";UNTIL=" + converterDatas(evento.data_termino, true) + "Z\n";
  texto = texto + "DTEND;TZID=America/Sao_Paulo:" + converterDatas(primeiraAula(evento)) + "T" + converterHorarios(evento.horario_termino) + "\n";
  texto = texto + "SUMMARY:" + evento.materia + "\n";
  texto = texto + "DESCRIPTION:Aula\n";
  texto = texto + "LOCATION:" + evento.local + "\n";
  texto = texto + "END:VEVENT\n";
  return texto;
}

function gerarICS()
{
  /*
      Gera o texto que vai entrar no arquivo ICS
  */
  let texto = "";
  texto = texto + "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//linkdosite\nCALSCALE:GREGORIAN\nBEGIN:VTIMEZONE\n";
  texto = texto + "TZID:America/Sao_Paulo\nTZURL:https://www.tzurl.org/zoneinfo-outlook/America/Sao_Paulo\n";
  texto = texto + "X-LIC-LOCATION:America/Sao_Paulo\nBEGIN:STANDARD\nTZNAME:-03\nTZOFFSETFROM:-0300\n";
  texto = texto + "TZOFFSETTO:-0300\nDTSTART:19700101T000000\nEND:STANDARD\nEND:VTIMEZONE\n";

  for (let i = 0; i < horario_compactado.length; i++)
  {
      texto = texto + gerarTextoICSEvento(horario_compactado[i]);
  }

  texto = texto + "END:VCALENDAR";
  return texto;
}

function downloadFile(text)
{
  /*
      Faz o download de um texto como arquivo, inicialmente era para arquivos .txt mas aparentemente
      funcionou para o ICS também então não vou reclamar kkkkkkkkkkkk
  */
    const link = document.createElement("a");
    const content = document.querySelector("textarea").value;
    const file = new Blob([text],
    {
        type: 'text/plain'
    });
    link.href = URL.createObjectURL(file);
    link.download = "horario.ics";
    link.click();
    URL.revokeObjectURL(link.href);
}