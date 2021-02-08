var ss = SpreadsheetApp.getActive();
var db = ss.getSheetByName("DB");
var mensal = ss.getSheetByName("Mensal");
var extrato = ss.getSheetByName("Extrato");
var input = ss.getSheetByName("Input");

var ultimaID = db.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getValue();
var dadosDB = db.getRange(1, 1, ultimaID, 11).getValues();

var hojeDia = Utilities.formatDate(new Date(), "GMT-04:00", "dd/MM/yy");

// Colunas da base de dados
var idColunaDB = 1;
var dataColunaDB = 2;
var origemColunaDB = 3;
var tipoColunaDB = 4;
var grupoColunaDB = 5;
var contaColunaDB = 6;
var nomeColunaDB = 7;
var valorColunaDB = 8;
var descricaoColunaDB = 9;
var pagoColunaDB = 10;
var prestacoesColunaDB = 11;
var iCallColunaDB = 14;

// Colunas do relatório Mensal
var mesAtualCol = 5;

// Colunas do extrato Mensal
var idExt = 2;
var dataExt = 3;
var grupoExt = 4;
var contaExt = 5;
var nomeExt = 7;
var valorExt = 8;
var pagoExt = 9;
var prestExt = 10;
var descExt = 11;
var origExt = 13;

// Limpa todos os dados da interface de entrada
function limparInput() {

  input.getRange("C4:C5").clearContent();
  input.getRange("F4:F5").clearContent();
  input.getRange("C6:D6").clearContent();
  input.getRange("F6").clearContent();
  input.getRange("C7:F8").clearContent();
  input.getRange("C10").clearContent();

}
// Atualiza a data no nome da planilha conforme dia atual
function atualizaDia() {

              // Gerando carimbo de data e versão
              var nomeAntigo = SpreadsheetApp.getActiveSpreadsheet().getName();
              var nomeSemData = nomeAntigo.slice(11); // Fim da data é na posição 10
              var novoNome = "[" + hojeDia + "] " + nomeSemData;

              SpreadsheetApp.getActiveSpreadsheet().rename(novoNome);

}

// Atualiza as informações da versão da ferramenta
function atualizaVersao() {

      // Define a linha da última versão
      var ultimaVersao = db.getRange(2, 20).getNextDataCell(SpreadsheetApp.Direction.DOWN).getValue();
      var lastRow = db.getRange(2, 20).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

      // Define os ranges onde são informadas os carimbos de versões atuais na notação a1
      var versInput = input.getRange("E14:F14");
      var versMensal = mensal.getRange("K36:L36");
      var versExtrato = extrato.getRange("K3");

      // Realiza prompts para inserção dos dados nos devidos campos
      var ui = SpreadsheetApp.getUi();

      var response = ui.alert("Oba! Eu vou ser mesmo atualizado?", ui.ButtonSet.YES_NO);

          if (response == ui.Button.YES) {

              var novaVersao = ui.prompt("Atualização", "Qual vai ser o nome da minha nova versão?", ui.ButtonSet.OK_CANCEL).getResponseText();
              var descricaoVersao = ui.prompt("Atualização", "O que mudou em relação à versão anterior (" + ultimaVersao + ")?", ui.ButtonSet.OK_CANCEL).getResponseText();

              // Gerando carimbo de data e versão
              var hojeDia = Utilities.formatDate(new Date(), "GMT-04:00", "dd/MM/yy");
              var nomeAtual = "[" + hojeDia + "] " + "Wallet v" + novaVersao;
              var carimboVersao = "versão " + novaVersao + " [" + hojeDia + "]";

              // Atualizando os ranges com os novos valores e renomeia a ferramenta
              versInput.setValue(carimboVersao)
              versMensal.setValue(carimboVersao)
              versExtrato.setValue(carimboVersao)
              SpreadsheetApp.getActiveSpreadsheet().rename(nomeAtual)

              db.getRange(lastRow + 1, 19).setValue(hojeDia)
              db.getRange(lastRow + 1, 20).setValue(novaVersao)
              db.getRange(lastRow + 1, 21).setValue(descricaoVersao);

              ui.alert("Valeu por cuidar de mim, cara!");

          } else if (response == ui.Button.NO) {

          ui.alert("Poxa, que pena!");

          }


}

// Verifica a conta referente à despesa e retorna o nome do grupo de despesas ao qual se enquadra
function defineGrupo(conta) {

  var conta = input.getRange("F5").getValue();

  if ( (conta == 'Faculdade') || (conta == 'Água') || (conta == 'Luz') || (conta == 'Combustível') || (conta == 'Claro') || (conta == 'Raio Rastreadores') || (conta == 'Seguro de Vida') ) {

    var grupo = "Despesas Fixas";

  } else if ( (conta == 'Bemol') || (conta == 'Riachuelo') || (conta == 'Saúde') ) {

    var grupo = 'Despesas Variáveis';

  } else if (conta == 'Crédito Pessoal') {

     var grupo = 'Crédito Pessoal/Limite';

  } else if  (conta == 'Crédito Compras') {

    var grupo = 'Cartão de Crédito';

  } else if ( (conta == 'Supermercado') || (conta == 'Refeição') || (conta == 'Lanches') ) {

    var grupo = 'Supermercado/Lanches';

  } else if ( (conta == 'Manutenção Carro') || (conta == 'Manutenção Outros') || (conta == 'Estacionamento') || (conta == 'Tarifas/Impostos') || (conta == 'Lazer') || (conta == 'Assinaturas') ) {

    var grupo = 'Despesas Ocasionais';

  } else if ( (conta == 'Veterinário') || (conta == 'Prejuízo Carro') || (conta == 'Oi') ) {

    var grupo = 'Imprevistos/Prejuízos';

  } else if ( (conta == 'Salário') || (conta == 'Rendas Diversas') ) {

    var grupo = 'Rendas';

  } else

    var grupo = 'Outros';

    return grupo;
}

// Converte mês em número (0 a 11)
function mesNumeroConv(mes) {

  if (mes == 'Janeiro') { return 0;
  } else if (mes == 'Fevereiro') { return 1;
    } else if (mes == 'Março') { return 2;
      } else if (mes == 'Abril') { return 3;
        } else if (mes == 'Maio') { return 4;
          } else if (mes == 'Junho') { return 5;
            } else if (mes == 'Julho') { return 6;
              } else if (mes == 'Agosto') { return 7;
                } else if (mes == 'Setembro') { return 8;
                  } else if (mes == 'Outubro') { return 9;
                    } else if (mes == 'Novembro') { return 10;
                      } else if (mes == 'Dezembro') { return 11;
                        } else ;


}
// Converte VERDADEIRO para 'SIM' e FALSO para 'NÃO'
function converteVerdadeiro(pago) {

  if (pago == 'VERDADEIRO') {
    return 'Sim';
  } else if (pago == 'FALSO') {
    return 'Não';
  } else return "?";

}

// Gera a lista dos registros de um determinado mês
function extratoMensal() {

  var mesNum = extrato.getRange(3, 10).getValue();
  var anoNum = extrato.getRange(2, 11).getValue();

  // Linha inicial da geração do extrato, contador de resultados obtidos (j)
  var linha = 7;
  var j = 0;

  // Limpa dados de buscas anteriores e retorna à formatação original
  extrato.getRange(7, 2, 100, 12).setBorder(false, false, false, false, false, false).clearContent().setFontWeight("normal").removeCheckboxes();

  dadosDB.forEach(function(row, i) {

  var id = row[0];
  var data = row[1];
  var origem = row[2];
  var tipo = row[3];
  var grupo = row[4];
  var conta = row[5];
  var nome = row[6];
  var valor = row[7];
  var descricao = row[8];
  var pago = row[9];
  var prestacoes = row[10];
  var mes = db.getRange(i + 1, 12).getValue();
  var ano = db.getRange(i + 1, 13).getValue();

        if ( (mesNum == mes) && (anoNum == ano) ) {

                extrato.getRange(linha + j, idExt).setValue(id);
                extrato.getRange(linha + j, dataExt).setValue(data);
                extrato.getRange(linha + j, grupoExt).setValue(grupo);
                extrato.getRange(linha + j, contaExt).setValue(conta);
                extrato.getRange(linha + j, nomeExt).setValue(nome);

                Logger.log("sim");

                          // Negativa valores caso seja do tipo 'Saída'; formata linha referente à crédito em conta
                          if (tipo == 'Entrada') {

                          extrato.getRange(linha + j, valorExt).setValue(valor);
                          extrato.getRange(linha + j, 2, 1, 12).setFontWeight("bold").setBorder(false, false, true, false, false, true);

                         } else extrato.getRange(linha + j, valorExt).setValue( (valor * (-1) ) );

                extrato.getRange(linha + j, pagoExt).setValue(pago).insertCheckboxes();
                extrato.getRange(linha + j, prestExt).setValue(prestacoes);
                extrato.getRange(linha + j, descExt).setValue(descricao);
                extrato.getRange(linha + j, origExt).setValue(origem);

                j++;

                } else Logger.log("não");

  });

            // Informa o saldo do período
            extrato.getRange(linha + j + 2, 7).setValue("SALDO DO PERÍODO:").setFontWeight("bold");
            extrato.getRange(linha + j + 2, 8).setFormula( ("=SUM(H7:H" + (linha + j) + ")") ).setFontWeight("bold");

}

// Insere dados informados na interface de entrada na base de dados (DB)
function enviar() {

  var ui = SpreadsheetApp.getUi();

  //Coletando ranges dos campos preenchidos
  var dataRange = input.getRange("C4");
  var origemRange = input.getRange("F4");
  var tipoRange = input.getRange("C5");
  var contaRange = input.getRange("F5");
  var nomeRange = input.getRange("C6:D6");
  var valorRange = input.getRange("F6");
  var descricaoRange = input.getRange("C7:F8");
  var pagoRange = input.getRange("C10");

  //Verificação de dados obrigatórios
  if ( (dataRange.isBlank()) || (origemRange.isBlank()) || (tipoRange.isBlank()) || (contaRange.isBlank()) || (nomeRange.isBlank()) || (valorRange.isBlank()) ) {

    ui.alert("Dados incompletos!");

  } else //if da verificação de dados obrigatórios

    var resposta = ui.alert("Confirma?", ui.ButtonSet.YES_NO);

    // Obtendo registro individual (código chave)
    var ultimaID = db.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
    var id = ultimaID + 1;
    //Logger.log(id);

    // Transferindo dados para a base
    if (resposta == ui.Button.YES) {

      var data = dataRange.getValue();
      var origem = origemRange.getValue();
      var tipo = tipoRange.getValue();
      var conta = contaRange.getValue();
      var nome = nomeRange.getValue();
      var valor = valorRange.getValue();
      var descricao = descricaoRange.getValue();
      var pago = pagoRange.isChecked();

      //Logger.log(origem);

      // Caso seja cartão de crédito, abre um prompt para a quantidade de prestações. O valor deverá gerar a quantidade de registros referente ao prazo de pagamento
      if ( (origem == 'Crédito next') || (origem == 'Crédito Nubank') || (origem == 'Crédito Renner') || (conta == 'Bemol') ) {

        var prestacoes = parseInt(ui.prompt(conta, "Informe a quantidade de prestações (0 se for à vista)" , ui.ButtonSet.OK_CANCEL).getResponseText());

         if (prestacoes == 0) {

                      prestacoes++;

                  } else if (prestacoes > 0) {

                            data.setMonth(data.getMonth() + 1);

                  } else ;

            for (i = 1; i <= prestacoes; i++) {

               // Transferindo os dados conforme registro
               db.getRange(id, idColunaDB).setValue(id);
               db.getRange(id, dataColunaDB).setValue(data);
               db.getRange(id, origemColunaDB).setValue(origem);
               db.getRange(id, tipoColunaDB).setValue(tipo);
               db.getRange(id, contaColunaDB).setValue(conta);
               db.getRange(id, nomeColunaDB).setValue(nome);
               db.getRange(id, valorColunaDB).setValue( (valor / prestacoes) );
               db.getRange(id, descricaoColunaDB).setValue(descricao);
               db.getRange(id, pagoColunaDB).setValue(pago).insertCheckboxes();
               db.getRange(id, prestacoesColunaDB).setValue(i + '/' + prestacoes);
               criarLembrete(id);

              // Classificando a entrada nos grupos definidos e inserindo na coluna ao qual pertence
              var grupo = defineGrupo(conta);
              db.getRange(id, grupoColunaDB).setValue(grupo);

              id++;
              data.setMonth(data.getMonth() + 1);

        }

      } else if ( (conta == 'Crédito Pessoal') && (tipo == 'Saída') ){

      var prestacoes = parseInt(ui.prompt("Crédito Pessoal", "Caso tenha sido parcelado, informe quantas prestações restam (0 se for à vista):" , ui.ButtonSet.OK_CANCEL).getResponseText());

                  if (prestacoes == 0) {

                      prestacoes++;

                  } else if (prestacoes > 0) {

                            data.setMonth(data.getMonth() + 1);

                  } else ;

            for (i = 1; i <= prestacoes; i++) {

               // Transferindo os dados conforme registro
               db.getRange(id, idColunaDB).setValue(id);
               db.getRange(id, dataColunaDB).setValue(data);
               db.getRange(id, origemColunaDB).setValue(origem);
               db.getRange(id, tipoColunaDB).setValue(tipo);
               db.getRange(id, contaColunaDB).setValue(conta);
               db.getRange(id, nomeColunaDB).setValue(nome);
               db.getRange(id, valorColunaDB).setValue( (valor) );
               db.getRange(id, descricaoColunaDB).setValue(descricao);
               db.getRange(id, pagoColunaDB).setValue(pago).insertCheckboxes();
               db.getRange(id, prestacoesColunaDB).setValue(i + '/' + prestacoes);
               criarLembrete(id);

              // Classificando a entrada nos grupos definidos e inserindo na coluna ao qual pertence
              var grupo = defineGrupo(conta);
              db.getRange(id, grupoColunaDB).setValue(grupo);

              id++;
              data.setMonth(data.getMonth() + 1);

              };


      } else if ( (origem !== 'Crédito next') && (origem !== 'Crédito Nubank') && (origem !== 'Crédito Renner') ) { // else da verificação de prestações

               // Transferindo os dados conforme registro
               db.getRange(id, idColunaDB).setValue(id);
               db.getRange(id, dataColunaDB).setValue(data);
               db.getRange(id, origemColunaDB).setValue(origem);
               db.getRange(id, tipoColunaDB).setValue(tipo);
               db.getRange(id, contaColunaDB).setValue(conta);
               db.getRange(id, nomeColunaDB).setValue(nome);
               db.getRange(id, valorColunaDB).setValue(valor);
               db.getRange(id, descricaoColunaDB).setValue(descricao);
               db.getRange(id, pagoColunaDB).setValue(pago).insertCheckboxes();
               db.getRange(id, prestacoesColunaDB).setValue("1");
               criarLembrete(id);

              // Classificando a entrada nos grupos definidos e inserindo na coluna ao qual pertence
              var grupo = defineGrupo(conta);
              db.getRange(id, grupoColunaDB).setValue(grupo);

      } else; // else do if não é de crédito

      atualizaDia();
      limparInput();

      // Atualiza a data da planilha assim que for enviado um novo registro



      } else ; // else do botão 'YES'

      input.getRange("C10").insertCheckboxes().check();
      dataRange.activate();
}
/*
// Gera o relatório RESUMO MENSAL
function resumoMensal() {

  // Localiza o período a ser gerado o relatório
  var mes = mensal.getRange("K2").getValue();
  var ano = mensal.getRange("L2").getValue();

  var mesNumero = mesNumeroConv(mes);

  // Preenchendo valores do campo Receitas - Despesas
  receitasDespesasAtual(mesNumero, ano);
  mesAtualCol--;
  receitasDespesasAtual(mesNumero - 1, ano);

  mesAtualCol = 5;

  }
*/
/****************************************************************/
// Inclui despesas não marcadas como PAGAS na soma de despesas do referido mês. O ID de entrada refere-se ao grupo em questão
// [0] Fixas [1] Variáveis [2] Cartão Crédito [3] Crédito Pessoal [4] Supermercado [5] Ocasionais [6] Imprevistos [7] Outros
function incNaoPagas (id) { // id

/*

    PAREI AQUI EM 07/02
    >>> TRANSFORMAR CÉLULAS AO CLICAR
    >>> FUNÇÃO NÃO É CHAMADA AUTOMATICAMENTE AO MUDAR O PERÍODO, VERIFICAR

*/

var mes = mensal.getRange("K3").getValue() - 1;
var ano = mensal.getRange("L2").getValue();
var mesAno = Utilities.formatDate(new Date(ano, mes), "GMT-04:00", "MM/yyyy");

Logger.log(mesAno + " mesAno");
var atual = verMesAtual(mesAno);

if (atual) {

var atrasados = buscarAtrasados();
var grupoAgg = agregarDespGrupos(atrasados);
Logger.log(atual  + " atual");
Logger.log(grupoAgg);

return grupoAgg[id];

} else return 0;


}

/****************************************************************/
// Atualiza as fórmulas do mês atual
// [0] Fixas [1] Variáveis [2] Cartão Crédito [3] Crédito Pessoal [4] Supermercado [5] Ocasionais [6] Imprevistos [7] Outros
function atualizarFormAutom () {

var formulas = [
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; B15; DB!D:D; "Saída") + incNaoPagas(0)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; B16; DB!D:D; "Saída") + incNaoPagas(1)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; B17; DB!D:D; "Saída") + incNaoPagas(2)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; B18; DB!D:D; "Saída") + incNaoPagas(3)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; "Supermercado/Lanches" ; DB!D:D; "Saída") + incNaoPagas(4)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; B20; DB!D:D; "Saída") + incNaoPagas(5)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; B21; DB!D:D; "Saída") + incNaoPagas(6)',
  '=SUMIFS(DB!H:H;DB!L:L; K3; DB!M:M; L2; DB!E:E; "Outros" ;DB!D:D; "Saída") + incNaoPagas(7)'
]

formulas.forEach(function (item, i) {

  mensal.getRange(15 + i, mesAtualCol).setValue(item);

});

}

/****************************************************************/
// Busca na base de dados as linhas de entrada e retorna um array com os valores agregados por grupo
// [0] Fixas [1] Variáveis [2] Cartão Crédito [3] Crédito Pessoal [4] Supermercado [5] Ocasionais [6] Imprevistos [7] Outros
function agregarDespGrupos (array) {

var lastRow = db.getRange(2, 10).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

var linhas = array;

var fixas = 0;
var variaveis = 0;
var cartao = 0;
var credPes = 0;
var mercado = 0;
var ocasion = 0;
var imprev = 0;
var outros = 0;

    linhas.forEach(function (item) {

    var id = db.getRange(item, 17).getValue();

    var grupo = db.getRange(id, grupoColunaDB).getValue();
    var valor = db.getRange(id, valorColunaDB).getValue();
    //var nome = db.getRange(id, nomeColunaDB).getValue();

        if (grupo == "Despesas Fixas") {
          fixas = fixas + valor;
        } else if (grupo == "Despesas Variáveis") {
          variaveis = variaveis + valor;
        } else if (grupo == "Cartão de Crédito") {
          cartao = cartao + valor;
        } else if (grupo == "Crédito Pessoal/Limite"){
          credPes = credPes + valor;
        } else if (grupo == "Supermercado/Lanches") {
          mercado = mercado + valor;
        } else if (grupo == "Despesas Ocasionais") {
          ocasion = ocasion + valor;
        } else if (grupo == "Imprevistos/Prejuízos") {
          imprev = imprev + valor;
        } else if (grupo == "Outros") {
          outros = outros + valor;
        }

    });

return [fixas, variaveis, cartao, credPes, mercado, ocasion, imprev, outros];

}
/****************************************************************/
// Busca dentro do DB de Contas não Pagas as linhas cuja data referenciada for menor que o dia atual. Retorna um array com as linhas
function buscarAtrasados(){

var linha = 33;
var coluna = 16;
var atrasados = [];

var hoje = parseInt(Utilities.formatDate(new Date(), "GMT-04:00", "D"));

var ult = db.getRange(linha, coluna).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
/*
var dataD = parseInt(Utilities.formatDate(new Date(data), "GMT-04:00", "D"));
var hojeD = parseInt(Utilities.formatDate(new Date(), "GMT-04:00", "D"));
var diasAtraso = hojeD - dataD;


    */

for (i = linha; i <= ult; i++) {

  var data = db.getRange(i, coluna).getValue();
  var dataAno = parseInt(Utilities.formatDate(data, "GMT-04:00", "yyyy"));
  var dataD = parseInt(Utilities.formatDate(data, "GMT-04:00", "D"));

    if (dataAno < 2021) {

      dataD = dataD - 365;

    }

    if (hoje > dataD) {
      atrasados.push(i);
    }


}

Logger.log(atrasados.length + " contas em atraso.");
return atrasados;

}

/****************************************************************/
// Retorna TRUE or FALSE se o mês atual for igual aos dados de entrada
function verMesAtual (mesAnoInput) {

var mesHoje = Utilities.formatDate(new Date(), "GMT-04:00", "MM/yyyy");
//var mesInp = mes + "/" + ano;
Logger.log(mesHoje);

if (mesHoje == mesAnoInput) {

  return true;

} else {

  return false;

}

}

/****************************************************************/
// Busca na base de dados os registros que constam como não pagos e preenche os dados no DB no range indicado
function verLinhasNaoPagas () {

var antiga = 132; // a linha mais antiga que retorna FALSE para pago (evitar calculos desnecessários);

var linhaDB = 33;
var colDataDB = 16;
var colLinhaDB = 17;

limparNaoPagos(linhaDB, colDataDB);

var lastRow = db.getRange(2, 10).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
var naoPagas = [];

  for (i = antiga; i <= lastRow; i++) {

    var range = db.getRange(i, pagoColunaDB);
    var pago = range.getValue();
    //Logger.log(typeof teste);

      if (!pago) {
        naoPagas.push(range.getRow());
      }

  }

  naoPagas.forEach(function (linha, i) {

    var data = db.getRange(linha, dataColunaDB).getValue();
    //Logger.log(i + ": " + item);
    db.getRange(linhaDB + i, colDataDB).setValue(data);
    db.getRange(linhaDB + i, colLinhaDB).setValue(linha);

    //Logger.log(data + " - linha: " + linha);

  });
  /*
  for (j = 0; j < naoPagas.length; j++) {


  }
  return naoPagas;
*/
Logger.log("sucesso");
}

/****************************************************************/
// Limpa os valores no DB de contas não pagas (indicadas no input)
function limparNaoPagos() { // linha, coluna

var linha = 33;
var coluna = 16;

var ult = db.getRange(linha, coluna).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

db.getRange(linha, coluna, ult - linha + 1, 2).clearContent();


}


/****************************************************************/
// Atualiza a fórmula de determinadas células (MANUAL)
function atualizaFormula() {

  var linha = 7;
  var coluna = 11;

  while (linha <= 33) {

   var range = String(mensal.getRange(linha, coluna).getFormula());
   var parte1 = range.slice(0, 26); // index é 26
   var parte2 = " DB!M:M; L2; ";
   var parte3 = range.slice(27);

   var novaFormula = parte1 + parte2 + parte3;
   mensal.getRange(linha, coluna).setFormula(novaFormula);
   linha++;

  }

}
