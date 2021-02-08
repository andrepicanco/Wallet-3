/***********************************************************************/

//                              REMINDERS                              //

/***********************************************************************/

// Any type of reminder will be integrated to this file
// Add-on to version 3.2.1

var email = "andreeeepicanco@gmail.com";
var agenda = "r082m3hb3dffe9l36slclln5qc@group.calendar.google.com";

/*###########################################################################*/

// Retorna a quantidade de registros atuais no banco de dados
function totalRegDB(){

var total = db.getRange(2, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getValue();

return total;

}

/*###########################################################################*/

// Verifica se a a ID em questão foi paga ou não, retorna TRUE ou FALSE
function checarPago(id) {

var pago = db.getRange(id, pagoColunaDB).getValue();

return pago;

}

/*###########################################################################*/

// Irá verificar cada id na base de dados sem pagamento, para criar um lembrete na agenda
function gerarLembretesFull() {

var registros = totalRegDB();

  for (i = 1; i <= registros; i++) {

    //criarLembrete(i);
    atualizarLembrete(i);

  }
    Logger.log("sucesso!");

}

/*###########################################################################*/
// Analisa os bugs em IDs localizados
function testarLembretes() {

var ids = []; // inserir IDs ao realizar testes

    for (i = 0; i < ids.length; i++) {

      Logger.log(ids[i]);
      //corrigirLembrete(ids[i]);
      criarLembrete(ids[i]);

    }

}

/*###########################################################################*/

// Cria um lembrete no Calendar do email caso a ID verificada não tiver sido paga
function criarLembrete(id) {

  var pago = checarPago(id);
  var hoje = new Date(); // devo alterar a data para o dia de vencimento do boleto

  if (!pago) {

    var dadosConta = buscarConta(id);
    var data = dadosConta[0];
    var conta = dadosConta[1];
    var nome = dadosConta[2];
    var valor = dadosConta[3];

    var dataForm = Utilities.formatDate(data, "GMT-04:00", "dd/MM/yy");
    var descricao = nome + " - R$ " + valor + "\nVencimento em: " + dataForm;

    var evento = CalendarApp.getCalendarById(agenda).createAllDayEvent(conta + " - " + nome, data,
            {description: descricao});

      var iCall = evento.getId();
      db.getRange(id, iCallColunaDB).setValue(iCall);

      var nDescricao = idDescricao(iCall);
      evento.setDescription(nDescricao);

  } else {



  }

}


/*###########################################################################*/

// Retorna informações de Data, Conta, Nome, Valor do pagamento e iCall referente à ID pesquisada em uma array
function buscarConta(id){

var data = db.getRange(id, dataColunaDB).getValue();
var conta = db.getRange(id, contaColunaDB).getValue();
var nome = db.getRange(id, nomeColunaDB).getValue();
var valor = db.getRange(id, valorColunaDB).getValue();
var obs = db.getRange(id, descricaoColunaDB).getValue();
var iCall = db.getRange(id, iCallColunaDB).getValue();

var dadosConta = [data, conta, nome, valor, obs, iCall];
return dadosConta;


}
/*###########################################################################*/

// Retorna a descrição atual de um determinado ID de evento, adicionando a ID no final da descrição
function idDescricao(id) {

var descricao = CalendarApp.getCalendarById(agenda).getEventById(id).getDescription();

var nDescricao = descricao.concat("\nID: " + id);

return nDescricao;

}

/*###########################################################################*/

// Busca na base de dados o código de iCall atual de calendário do registro
function buscarIcall(id) {

var iCall = db.getRange(id, iCallColunaDB).getValue();

return iCall;

}

/*###########################################################################*/

// Atualiza para o dia atual a id de pagamento que ainda não tiver sido paga
function atualizarLembrete(id) {

var pago = checarPago(id);

var dadosConta = buscarConta(id);
var data = dadosConta[0];
var conta = dadosConta[1];
var nome = dadosConta[2];
var valor = dadosConta[3];
var iCall = dadosConta[4];

var dataAno = Utilities.formatDate(new Date(data), "GMT-04:00", "yyyy");
var dataD = parseInt(Utilities.formatDate(new Date(data), "GMT-04:00", "D"));
var hojeD = parseInt(Utilities.formatDate(new Date(), "GMT-04:00", "D"));
var diasAtraso = hojeD - dataD;

    if (dataAno < 2021) {

      dataD = dataD - 365;
      diasAtraso = (dataD * (-1)) + parseInt(hojeD);

    }

var hoje = new Date();

    if ( (!pago) && (hojeD > dataD) ) {

      var iCallAntigo = buscarIcall(id);
      CalendarApp.getCalendarById(agenda).getEventById(iCallAntigo).deleteEvent();

      var dataForm = Utilities.formatDate(data, "GMT-04:00", "dd/MM/yy");
      var descricao = nome + " - R$ " + valor + "\nVencido em: " + dataForm
                      + " (" + diasAtraso + " dias em atraso)";

      var novoEvento = CalendarApp.getCalendarById(agenda).createAllDayEvent(conta + " - " + nome, hoje,
            {description: descricao});

      var idNovoEvento = novoEvento.getId();
      db.getRange(id, iCallColunaDB).setValue(idNovoEvento);

      var nDescricao = idDescricao(idNovoEvento);
      novoEvento.setDescription(nDescricao);


    } else {

      Logger.log("ainda não venceu ou foi pago");

    }

}
