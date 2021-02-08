
/**********************************************************************************************************************************************************************************/

                  /*                                                                Módulo UBER                                                                */

/**********************************************************************************************************************************************************************************/

                                  // VARIÁVEIS GLOBAIS
var db = ss.getSheetByName("DB");
var ss = SpreadsheetApp.getActive();
var uber = ss.getSheetByName("Uber");
var lastRow = db.getRange(1, 23).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

// Ranges da aba de INPUT
var data = uber.getRange("C4");
var tempo = uber.getRange("C5");
var viagens = uber.getRange("E4");
var distancia = uber.getRange("E5");
var fatorKM = uber.getRange("C6");
var bruto = uber.getRange("E6");
var rentab = uber.getRange("C7");
var liquido = uber.getRange("E7");

// Colunas da base de dados UBER
var dataDBU = 23;
var tempoDBU = 24;
var viagensDBU = 25;
var distanciaDBU = 26;
var fatorKMDBU = 27;
var brutoDBU = 28;
var rentabDBU = 29;
var liquidoDBU = 30;

// KPIs Uber
var semana = 33;
var lucro = 34;
var rentabSem = 35;
var spreadKM = 36;
var viagensDia = 37;
var lucroViag = 38;
var tempoDia = 39;

// Limpa campos do input
function limparUber() {

data.clearContent()
tempo.clearContent()
viagens.clearContent()
distancia.clearContent()
fatorKM.clearContent()
bruto.clearContent();

}

// Envia dados do input para a base de dados
function enviarUber() {

if (  (data.isBlank()) || (tempo.isBlank()) || (viagens.isBlank()) || (distancia.isBlank()) || (fatorKM.isBlank()) || (bruto.isBlank()) || (rentab.isBlank()) || (liquido.isBlank()) ) {

  SpreadsheetApp.getUi().alert("Opa irmão, tá faltando coisa aí hein!");

    } else {

      db.getRange(lastRow + 1, dataDBU).setValue(data.getValue())
      db.getRange(lastRow + 1, tempoDBU).setValue(tempo.getValue())
      db.getRange(lastRow + 1, viagensDBU).setValue(viagens.getValue())
      db.getRange(lastRow + 1, distanciaDBU).setValue(distancia.getValue())
      db.getRange(lastRow + 1, fatorKMDBU).setValue(fatorKM.getValue())
      db.getRange(lastRow + 1, brutoDBU).setValue(bruto.getValue())
      db.getRange(lastRow + 1, rentabDBU).setValue(rentab.getValue())
      db.getRange(lastRow + 1, liquidoDBU).setValue(liquido.getValue());

      limparUber();

        }

}

// Retorna a data da última entrada de dados
function ultimaEntrada() {

var ultEntrada = db.getRange(1, 23).getNextDataCell(SpreadsheetApp.Direction.DOWN).getValue();

return ultEntrada;

}

// Preenche semana e mês atual
function atual() {

uber.getRange("M2").setFormula("=weeknum(today())")
uber.getRange("Q2").setFormula("=MONTH(today())");


}

// Exporta KPIs para base de dados
function exportarKPIs() {

var lastWeek = db.getRange(1, 33).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

var semanaKPI = uber.getRange("M2").getValue();
var lucroKPI = uber.getRange("J4").getValue();
var rentabSemKPI = uber.getRange("J6").getValue();
var spreadKMKPI = uber.getRange("J9").getValue();
var viagensDiaKPI = uber.getRange("J12").getValue();
var lucroViagKPI = uber.getRange("J15").getValue();
var tempoDiaKPI = uber.getRange("J16").getValue();

db.getRange(lastWeek + 1, semana).setValue(semanaKPI);
db.getRange(lastWeek + 1, lucro).setValue(lucroKPI);
db.getRange(lastWeek + 1, rentabSem).setValue(rentabSemKPI);
db.getRange(lastWeek + 1, spreadKM).setValue(spreadKMKPI);
db.getRange(lastWeek + 1, viagensDia).setValue(viagensDiaKPI);
db.getRange(lastWeek + 1, lucroViag).setValue(lucroViagKPI);
db.getRange(lastWeek + 1, tempoDia).setValue(tempoDiaKPI);

SpreadsheetApp.getUi().alert("show, mano!");

}
