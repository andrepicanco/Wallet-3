/***********************************************************************/

//                                PLANO                                //

/***********************************************************************/

// Scripts used for Plan Sheet
// Add-on to version 3.2.1

// GLOBAIS
var plano = ss.getSheetByName("Plano");

var tri1 = ["Janeiro", "Fevereiro", "Março"];
var tri2 = ["Abril", "Maio", "Junho"];
var tri3 = ["Julho", "Agosto", "Setembro"];
var tri4 = ["Outubro", "Novembro", "Dezembro"];

// RANGES IMPORTANTES
var triBusca = "K2";
var anoBusca = "L2";
var mesZeroLinha = [4, 6, 19] // Localização das linhas onde ficam o mês 0 na interface
var mesZeroCol = [3, 10, 10]; // Localização das colunas onde ficam o mês 0 na interface

var objetivo = "C23";
var obs = "B25";

var receitasGeral = ["C5", "D5", "E5", "F5"]; // Mês [0], Mês [1], Mês[2] e Mês[3]
var despesasGeral = ["C6", "D6", "E6", "F6"]; // Mês [0], Mês [1], Mês[2] e Mês[3]

// RANGES: DESPESA PROJETADA (Mês [0], Mês [1], Mês[2] e Mês[3])
var despFixas = ["J7", "K7", "L7", "M7"];
var despVar = ["J8", "K8", "L8", "M8"];
var despCart = ["J9", "K9", "L9", "M9"];
var despCred = ["J10", "K10", "L10", "M10"];
var despMerc = ["J11", "K11", "L11", "M11"];
var despOcas = ["J12", "K12", "L12", "M12"];
var despImprev = ["J13", "K13", "L13", "M13"];
var desOutro  = ["J14", "K14", "L14", "M14"];

// RANGES: RECEITA PROJETADA (Mês [0], Mês [1], Mês[2] e Mês[3])
var recSal = ["J20", "K20", "L20", "M20"];
var recAdic = ["J21", "K21", "L21", "M21"];
var recCred = ["J22", "K22", "L22", "M22"];
var recOutro = ["J23", "K23", "L23", "M23"];
var recExtra = ["J24", "K24", "L24", "M24"];

// RANGES: TAREFAS ([0] Nome, [1] CheckBox, [2] Descrição)
var tar1 = ["I30", "J30", "K30"];
var tar2 = ["I31", "J31", "K31"];
var tar3 = ["I32", "J32", "K32"];
var tar4 = ["I33", "J33", "K33"];
var tar5 = ["I34", "J34", "K34"];

/*#############################################################################################################*/
// Define os nomes nos ranges de mês dependendo do período selecionado
function nomearPeriodos(){

var tri = plano.getRange(triBusca).getValue();
var ano = plano.getRange(anoBusca).getValue();

    if (tri == '1T') {

        for (i = 0; i < mesZeroCol.length; i++) {

            plano.getRange(mesZeroLinha[i], mesZeroCol[i]).setValue("Dezembro/" + parseInt(ano - 1));
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 1).setValue(tri1[0] + "/" + ano);
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 2).setValue(tri1[1] + "/" + ano);
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 3).setValue(tri1[2] + "/" + ano);

        }

    } else if (tri == '2T') {

        for (i = 0; i < mesZeroCol.length; i++) {

            plano.getRange(mesZeroLinha[i], mesZeroCol[i]).setValue("Março/" + parseInt(ano));
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 1).setValue(tri2[0] + "/" + ano);
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 2).setValue(tri2[1] + "/" + ano);
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 3).setValue(tri2[2] + "/" + ano);

        }

    } else if (tri == '3T') {

        for (i = 0; i < mesZeroCol.length; i++) {

            plano.getRange(mesZeroLinha[i], mesZeroCol[i]).setValue("Junho/" + parseInt(ano));
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 1).setValue(tri3[0] + "/" + ano);
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 2).setValue(tri3[1] + "/" + ano);
            plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 3).setValue(tri3[2] + "/" + ano);

        }

    } else if (tri == "4T") {

        for (i = 0; i < mesZeroCol.length; i++) {

          plano.getRange(mesZeroLinha[i], mesZeroCol[i]).setValue("Setembro/" + parseInt(ano));
          plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 1).setValue(tri4[0] + "/" + ano);
          plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 2).setValue(tri4[1] + "/" + ano);
          plano.getRange(mesZeroLinha[i], mesZeroCol[i] + 3).setValue(tri4[2] + "/" + ano);

      }

    }

}

/*#############################################################################################################*/
// Insere lista suspensa no range de Objetivos com os valores já cadastrados na base de dados
function listarObjetivos (){
}
