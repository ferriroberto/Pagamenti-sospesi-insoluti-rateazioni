////////////////////////////////////////////////////////////////////////////////////
/////////////////////pagamenti e sospesi////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Le intestazioni si trovano alla riga 2 (indice 1 nell'array)
  var headers = data[1];

  // Trova gli indici delle colonne necessarie
  var emailColumnIndex = headers.indexOf("mail");
  var clienteColumnIndex = headers.indexOf("CLIENTE");
  var descrizioneColumnIndex = headers.indexOf("DESCRIZIONE INSOLUTO/SOSPESO");
  var dataScadenzaColumnIndex = headers.indexOf("DATA SCADENZA");
  var saldoColumnIndex = headers.indexOf("SALDO");
  var azioneColumnIndex = headers.indexOf("Azione");
  var numInviiColumnIndex = headers.indexOf("num.invii");

  // Verifica che tutte le colonne richieste esistano
  if (emailColumnIndex === -1 || clienteColumnIndex === -1 ||
    descrizioneColumnIndex === -1 || saldoColumnIndex === -1 ||
    azioneColumnIndex === -1) {
    Browser.msgBox("Errore", "Una o più colonne richieste non sono state trovate. Verifica che esistano le colonne: mail, CLIENTE, DESCRIZIONE, SALDO, Azione", Browser.Buttons.OK);
    return;
  }

  // Verifica se esiste la colonna num.invii
  if (numInviiColumnIndex === -1) {
    Browser.msgBox("Errore", "Colonna 'num.invii' non trovata. Assicurati che esista questa colonna nella riga delle intestazioni.", Browser.Buttons.OK);
    return;
  }

  // Oggetto per raccogliere i dati per cliente
  var clienteMap = {};
  var rowiPerCliente = {};

  // Prima fase: raccogliere tutti gli insoluti per cliente
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var emailAddress = row[emailColumnIndex];
    var cliente = row[clienteColumnIndex];
    var descrizione = row[descrizioneColumnIndex];
    var dataGrezza = row[dataScadenzaColumnIndex];
    var saldo = row[saldoColumnIndex];
    var azione = row[azioneColumnIndex];
    
    // Verifica se l'azione è TRUE e se c'è un indirizzo email
    if (azione === true && emailAddress) {
      // Formatta il saldo
      var saldoFormattato;
      if (typeof saldo === 'number') {
        saldoFormattato = saldo.toFixed(2);
        saldoFormattato = saldoFormattato.replace('.', 'DECIMALE');
        saldoFormattato = saldoFormattato.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        saldoFormattato = saldoFormattato.replace('DECIMALE', ',');
      } else {
        saldoFormattato = saldo;
      }

      // Formatta la data
      var dataFormattata = "";
      if (dataGrezza instanceof Date) {
        dataFormattata = Utilities.formatDate(dataGrezza, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      // Crea una chiave univoca per cliente e email
      var chiave = cliente + "|||" + emailAddress;
      
      // Inizializza l'array se non esiste
      if (!clienteMap[chiave]) {
        clienteMap[chiave] = [];
        rowiPerCliente[chiave] = [];
      }
      
      // Aggiungi i dati di questo insoluto
      clienteMap[chiave].push({
        descrizione: descrizione,
        dataFormattata: dataFormattata,
        saldoFormattato: saldoFormattato
      });
      
      // Memorizza l'indice della riga per aggiornare num.invii dopo
      rowiPerCliente[chiave].push(i);
    }
  }

  var emailsInviati = 0;

  // Seconda fase: invia email consolidate
  for (var chiave in clienteMap) {
    var parti = chiave.split("|||");
    var cliente = parti[0];
    var emailAddress = parti[1];
    var insoluti = clienteMap[chiave];
    
    if (insoluti.length > 0) {
      // Crea il contenuto dell'email
      var subject = "Promemoria adempimenti - messaggio automatico";
      var message = "Gentile " + cliente + ",\n\n";
      message += "Con la presente desideriamo semplicemente ricordare che risultano ancora aperte le posizioni relative ";
      message += "ai seguenti adempimenti:\n\n";
      
      // Aggiungi ogni insoluto in formato elenco
      for (var j = 0; j < insoluti.length; j++) {
        var insoluto = insoluti[j];
        message += (j+1) + ") \"" + insoluto.descrizione + "\" scaduta in data " + insoluto.dataFormattata;
        message += " per un importo di € " + insoluto.saldoFormattato + "\n";
      }
      
      message += "\nNaturalmente, segnaliamo la cosa solo per agevolare una reciproca e serena gestione dei rapporti.\n";
      message += "Nel caso avesse già provveduto o avesse necessità di ulteriori informazioni, La invito a non esitare a contattarci. Nel caso, invece, ";
      message += "ricordava lo stato di sospensione di questi adempimenti e intende, al momento, non intervenire non dovrà fare nulla e può anche cestinare la presente.\n\n";
      message += "RingraziandoLa per l'attenzione e con l'occasione porgiamo i nostri più cordiali saluti.,\n\nElaborazionie e Servizi s.r.l.";

      // Invia l'email
      try {
        GmailApp.sendEmail(emailAddress, subject, message, {
          from: "info@elaborazionieservizi.com" // Alias configurato in Gmail
        });
        emailsInviati++;

        // Incrementa num.invii per tutte le righe di questo cliente
        var righe = rowiPerCliente[chiave];
        for (var k = 0; k < righe.length; k++) {
          var rowIndex = righe[k];
          var numInvii = data[rowIndex][numInviiColumnIndex] || 0;
          numInvii = (typeof numInvii === 'number') ? numInvii + 1 : 1;
          sheet.getRange(rowIndex + 1, numInviiColumnIndex + 1).setValue(numInvii);
        }

      } catch (error) {
        Logger.log("Errore nell'invio dell'email a " + emailAddress + ": " + error.message);
      }
    }
  }

  // Mostra un messaggio di conferma
  if (emailsInviati > 0) {
    Browser.msgBox("Successo", "Sono stati inviati " + emailsInviati + " promemoria consolidati.", Browser.Buttons.OK);
  } else {
    Browser.msgBox("Informazione", "Nessun promemoria da inviare. Verifica che ci siano righe con Azione = TRUE.", Browser.Buttons.OK);
  }
}

// Funzione per mostrare un popup di filtro e inviare email filtrate
function showFilterDialog() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Le intestazioni si trovano alla riga 2 (indice 1)
  var headers = data[1];

  var clienteColumnIndex = headers.indexOf("CLIENTE");
  var operatoreColumnIndex = headers.indexOf("operatore");

  if (clienteColumnIndex === -1) {
    Browser.msgBox("Errore", "Colonna 'CLIENTE' non trovata. Verifica che esista questa colonna.", Browser.Buttons.OK);
    return;
  }

  var operatoriPresenti = operatoreColumnIndex !== -1;
  var operatori = [];
  var clienti = [];

  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var cliente = row[clienteColumnIndex];
    if (cliente && clienti.indexOf(cliente) === -1) clienti.push(cliente);
    if (operatoriPresenti) {
      var operatore = row[operatoreColumnIndex];
      if (operatore && operatori.indexOf(operatore) === -1) operatori.push(operatore);
    }
  }

  clienti.sort();
  if (operatoriPresenti) operatori.sort();

  var html = '<html><head>'
    + '<style>'
    + 'body { font-family: Arial, sans-serif; margin: 0; padding: 20px; display: flex; justify-content: center; align-items: center; height: 90vh; }'
    + '.container { background: white; padding: 30px 40px; max-width: 800px; width: 100%; height: 80%}'
    + 'h2 { color: #333; text-align: center; margin-bottom: 20px; font-size: 22px; }'
    + 'p { color: #555; font-size: 14px; line-height: 1.5; margin-bottom: 20px; text-align: center; }'
    + 'label { display: block; margin: 20px 0 8px; color: #444; font-size: 14px; }'
    + 'select { width: 100%; height: 450px; padding: 10px; border: 1px solid #ccc; border-radius: 8px; font-size: 14px; transition: border-color 0.3s; }'
    + 'select:focus { outline: none; border-color: #007bff; }'
    + '.button-container { margin-top: 30px; display: flex; justify-content: space-between; }'
    + 'button { flex: 1; padding: 12px; margin: 0 5px; border: none; border-radius: 8px; background: #007bff; color: white; cursor: pointer; font-size: 15px; transition: background 0.3s, transform 0.1s; }'
    + 'button:hover { background: #0056b3; }'
    + 'button:active { transform: scale(0.98); }'
    + 'button.cancel { background: #e0e0e0; color: #333; }'
    + 'button.cancel:hover { background: #c7c7c7; }'
    + '</style>'
    + '</head><body>'
    + '<div class="container">'
    + '<h2>Filtra invio email</h2>'
    + '<p>Seleziona i criteri per filtrare le email. Saranno inviate solo quelle con "Azione" impostato su <b>TRUE</b>.</p>';


  if (operatoriPresenti && operatori.length > 0) {
    html += '<label for="operatore">Filtra per Operatore:</label>'
      + '<select id="operatore">'
      + '<option value="">-- Tutti gli operatori --</option>';
    for (var i = 0; i < operatori.length; i++) {
      html += '<option value="' + operatori[i] + '">' + operatori[i] + '</option>';
    }
    html += '</select>';
  }

  if (clienti.length > 0) {
    html += '<label for="cliente">Filtra per Cliente (puoi selezionare più clienti tenendo premuto CTRL o CMD):</label>'
      + '<select id="cliente" multiple size="6">'; // size per dare spazio visivo
    for (var i = 0; i < clienti.length; i++) {
      html += '<option value="' + clienti[i] + '">' + clienti[i] + '</option>';
    }
    html += '</select>';
  }

  html += '<div class="button-container">'
    + '<button class="cancel" onclick="google.script.host.close()">Annulla</button>'
    + '<button id="btnInvia" onclick="inviaEmailFiltrate()">Invia Email</button>'
    + '</div>'
    + '<script>'
    + 'function inviaEmailFiltrate() {'
    + '  var btn = document.getElementById("btnInvia");'
    + '  btn.disabled = true;' // Disabilita il pulsante
    + '  btn.innerHTML = "<span style=\'display: inline-flex; align-items: center;\'><div class=\'loader\'></div> Invia...</span>";' // Mostra loader'
    + ''
    + '  var operatore = document.getElementById("operatore") ? document.getElementById("operatore").value : "";'
    + '  var clienteSelect = document.getElementById("cliente");'
    + '  var clientiSelezionati = [];'
    + '  if (clienteSelect) {'
    + '    for (var i = 0; i < clienteSelect.options.length; i++) {'
    + '      if (clienteSelect.options[i].selected) {'
    + '        clientiSelezionati.push(clienteSelect.options[i].value);'
    + '      }'
    + '    }'
    + '  }'
    + ''
    // Quando l'invio è completato con successo, chiudi subito il dialog
    + '  google.script.run.withSuccessHandler(function() {'
    + '    google.script.host.close();' // Chiusura prima del messaggio di successo
    + '  }).inviaEmailConFiltro(operatore, clientiSelezionati);'
    + '}'
    + '</script>'
    + '<style>'
    + '.loader {'
    + '  border: 2px solid #f3f3f3;'
    + '  border-top: 2px solid #3498db;'
    + '  border-radius: 50%;'
    + '  width: 14px;'
    + '  height: 14px;'
    + '  animation: spin 1s linear infinite;'
    + '  margin-right: 5px;'
    + '}'
    + '@keyframes spin {'
    + '  0% { transform: rotate(0deg); }'
    + '  100% { transform: rotate(360deg); }'
    + '}'
    + '</style>'
    + '</body></html>';


  var ui = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Filtra invio email');
}

// Funzione per inviare email con filtri
function inviaEmailConFiltro(operatoreFiltro, clientiFiltroArray) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Le intestazioni si trovano alla riga 2 (indice 1 nell'array)
  var headers = data[1];

  // Trova gli indici delle colonne necessarie
  var emailColumnIndex = headers.indexOf("mail");
  var clienteColumnIndex = headers.indexOf("CLIENTE");
  var descrizioneColumnIndex = headers.indexOf("DESCRIZIONE INSOLUTO/SOSPESO");
  var dataScadenzaColumnIndex = headers.indexOf("DATA SCADENZA");
  var saldoColumnIndex = headers.indexOf("SALDO");
  var azioneColumnIndex = headers.indexOf("Azione");
  var numInviiColumnIndex = headers.indexOf("num.invii");
  var operatoreColumnIndex = headers.indexOf("operatore");

  // Verifica che tutte le colonne richieste esistano
  if (emailColumnIndex === -1 || clienteColumnIndex === -1 ||
    descrizioneColumnIndex === -1 || saldoColumnIndex === -1 ||
    azioneColumnIndex === -1) {
    Browser.msgBox("Errore", "Una o più colonne richieste non sono state trovate.", Browser.Buttons.OK);
    return;
  }

  // Verifica se esiste la colonna num.invii
  if (numInviiColumnIndex === -1) {
    Browser.msgBox("Errore", "Colonna 'num.invii' non trovata. Assicurati che esista questa colonna nella riga delle intestazioni.", Browser.Buttons.OK);
    return;
  }

  // Oggetto per raccogliere i dati per cliente
  var clienteMap = {};
  var rowiPerCliente = {};

  // Prima fase: raccogliere tutti gli insoluti per cliente/email che passano i filtri
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var emailAddress = row[emailColumnIndex];
    var cliente = row[clienteColumnIndex];
    var descrizione = row[descrizioneColumnIndex];
    var dataGrezza = row[dataScadenzaColumnIndex];
    var saldo = row[saldoColumnIndex];
    var azione = row[azioneColumnIndex];
    var operatore = operatoreColumnIndex !== -1 ? row[operatoreColumnIndex] : "";

    // Filtra per operatore (se selezionato) e cliente (se uno o più selezionati)
    var passaFiltroOperatore = !operatoreFiltro || operatore === operatoreFiltro;
    var passaFiltroCliente = clientiFiltroArray.length === 0 || clientiFiltroArray.indexOf(cliente) !== -1;

    // Procedi solo se filtri superati, Azione = TRUE e mail presente
    if (azione === true && emailAddress && passaFiltroOperatore && passaFiltroCliente) {
      // Formattazione saldo
      var saldoFormattato = (typeof saldo === 'number')
        ? saldo.toFixed(2).replace('.', 'DECIMALE').replace(/\B(?=(\d{3})+(?!\d))/g, ".").replace('DECIMALE', ',')
        : saldo;

      // Formattazione data scadenza
      var dataFormattata = (dataGrezza instanceof Date)
        ? Utilities.formatDate(dataGrezza, Session.getScriptTimeZone(), "dd/MM/yyyy")
        : "";

      // Crea una chiave univoca per cliente e email
      var chiave = cliente + "|||" + emailAddress;
      
      // Inizializza l'array se non esiste
      if (!clienteMap[chiave]) {
        clienteMap[chiave] = [];
        rowiPerCliente[chiave] = [];
      }
      
      // Aggiungi i dati di questo insoluto
      clienteMap[chiave].push({
        descrizione: descrizione,
        dataFormattata: dataFormattata,
        saldoFormattato: saldoFormattato
      });
      
      // Memorizza l'indice della riga per aggiornare num.invii dopo
      rowiPerCliente[chiave].push(i);
    }
  }

  var emailsInviati = 0;

  // Seconda fase: invia email consolidate
  for (var chiave in clienteMap) {
    var parti = chiave.split("|||");
    var cliente = parti[0];
    var emailAddress = parti[1];
    var insoluti = clienteMap[chiave];
    
    if (insoluti.length > 0) {
      // Crea il contenuto dell'email
      var subject = "Promemoria adempimenti - messaggio automatico";
      var message = "Gentile " + cliente + ",\n\n";
      message += "Con la presente desideriamo semplicemente ricordare che risultano ancora aperte le posizioni relative ";
      message += "ai seguenti adempimenti:\n\n";
      
      // Aggiungi ogni insoluto in formato elenco
      for (var j = 0; j < insoluti.length; j++) {
        var insoluto = insoluti[j];
        message += (j+1) + ") \"" + insoluto.descrizione + "\" scaduta in data " + insoluto.dataFormattata;
        message += " per un importo di € " + insoluto.saldoFormattato + "\n";
      }
      
      message += "\nNaturalmente, segnaliamo la cosa solo per agevolare una reciproca e serena gestione dei rapporti.\n";
      message += "Nel caso avesse già provveduto o avesse necessità di ulteriori informazioni, La invito a non esitare a contattarci. Nel caso, invece, ";
      message += "ricordava lo stato di sospensione di questi adempimenti e intende, al momento, non intervenire non dovrà fare nulla e può anche cestinare la presente.\n\n";
      message += "RingraziandoLa per l'attenzione e con l'occasione porgiamo i nostri più cordiali saluti.,\n\nElaborazionie e Servizi s.r.l.";

      // Invia l'email
      try {
        GmailApp.sendEmail(emailAddress, subject, message, {
          from: "info@elaborazionieservizi.com" // Alias configurato in Gmail
        });
        emailsInviati++;

        // Incrementa num.invii per tutte le righe di questo cliente
        var righe = rowiPerCliente[chiave];
        for (var k = 0; k < righe.length; k++) {
          var rowIndex = righe[k];
          var numInvii = data[rowIndex][numInviiColumnIndex] || 0;
          numInvii = (typeof numInvii === 'number') ? numInvii + 1 : 1;
          sheet.getRange(rowIndex + 1, numInviiColumnIndex + 1).setValue(numInvii);
        }

      } catch (error) {
        Logger.log("Errore nell'invio dell'email a " + emailAddress + ": " + error.message);
      }
    }
  }

  // Messaggio finale
  if (emailsInviati > 0) {
    Browser.msgBox("Successo", "Sono stati inviati " + emailsInviati + " promemoria consolidati.", Browser.Buttons.OK);
  } else {
    Browser.msgBox("Informazione", "Nessun promemoria da inviare. Controlla i filtri o le colonne Azione e Mail.", Browser.Buttons.OK);
  }
}

//////////////////////////////////////////////////////////////////////////////////////////////
/////////foglio Rateazioni///////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////
function creaScadenzeNelCalendario() {
  // Imposta il numero fisso di giorni prima per la notifica email
  var giorniPrima = 3;
  
  // Ottiene il foglio specifico "Rateazioni"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rateazioni");
  
  // Verifica che il foglio esista
  if (sheet == null) {
    SpreadsheetApp.getUi().alert("Il foglio 'Rateazioni' non è stato trovato!");
    return;
  }
  
  // Ottiene i dati dal foglio (escludendo l'intestazione)
  var dati = sheet.getDataRange().getValues();
  var intestazioni = dati[0]; // La prima riga contiene le intestazioni
  
  // Trova gli indici delle colonne rilevanti
  var indiceCliente = intestazioni.indexOf("CLIENTE");
  var indiceRata = intestazioni.indexOf("Rata");
  var indiceData = intestazioni.indexOf("Data");
  var indiceDescrizione = intestazioni.indexOf("Descrizione");
  var indiceDaVersare = intestazioni.indexOf("Da versare");
  
  // Verifica che tutte le colonne necessarie esistano
  if (indiceCliente < 0 || indiceRata < 0 || indiceData < 0 || 
      indiceDescrizione < 0 || indiceDaVersare < 0) {
    throw new Error("Una o più colonne richieste non sono state trovate nel foglio 'Rateazioni'.");
  }
  
  // Ottiene il calendario predefinito
  var calendario = CalendarApp.getDefaultCalendar();
  
  // Conta quanti eventi sono stati creati e quanti erano già presenti
  var contatoreCreazioneEventi = 0;
  var contatoreEventiEsistenti = 0;
  
  // Itera attraverso tutte le righe di dati (iniziando dalla seconda riga, escludendo le intestazioni)
  for (var i = 1; i < dati.length; i++) {
    var riga = dati[i];
    
    // Estrae i valori dalle celle
    var cliente = riga[indiceCliente];
    var rata = riga[indiceRata];
    var data = riga[indiceData];
    var descrizione = riga[indiceDescrizione];
    var importo = riga[indiceDaVersare];
    
    // Salta righe vuote
    if (!cliente && !importo) {
      continue;
    }
    
    // Verifica che la data sia valida
    if (data instanceof Date && !isNaN(data.getTime())) {
      // Crea il titolo dell'evento
      var titoloEvento = cliente + " - " + descrizione + " (Rata n. " + rata + ") - " + importo + "€";
      
      // Cerca se l'evento esiste già
      var inizio = new Date(data);
      inizio.setHours(0, 0, 0, 0); // Inizio della giornata
      
      var fine = new Date(data);
      fine.setHours(23, 59, 59, 999); // Fine della giornata
      
      var eventiEsistenti = calendario.getEvents(inizio, fine);
      var eventoEsistente = false;
      
      // Controlla se esiste già un evento con lo stesso titolo e nella stessa data
      for (var j = 0; j < eventiEsistenti.length; j++) {
        var titoloDaControllare = eventiEsistenti[j].getTitle();
        
        // Se il titolo corrisponde, consideriamo l'evento come già esistente
        if (titoloDaControllare === titoloEvento) {
          eventoEsistente = true;
          contatoreEventiEsistenti++;
          break;
        }
      }
      
      // Se l'evento non esiste già, lo creiamo
      if (!eventoEsistente) {
        // Crea una descrizione dettagliata per l'evento
        var descrizioneEvento = "Cliente: " + cliente + "\n" +
                                "Rata n.: " + rata + "\n" +
                                "Descrizione: " + descrizione + "\n" +
                                "Importo da versare: " + importo + "€";
        
        // Crea l'evento nel calendario
        var evento = calendario.createAllDayEvent(
          titoloEvento,
          data,
          {description: descrizioneEvento}
        );
        
        // Aggiunge la notifica email 3 giorni prima
        evento.addEmailReminder(giorniPrima * 24 * 60); // Converti giorni in minuti (3 giorni = 4320 minuti)
        
        contatoreCreazioneEventi++;
      }
    }
  }
  
  // Mostra un messaggio di completamento
  SpreadsheetApp.getUi().alert("Operazione completata!\n" +
                               "Nuovi eventi creati: " + contatoreCreazioneEventi + "\n" +
                               "Eventi già esistenti: " + contatoreEventiEsistenti + "\n\n" +
                               "Riceverai un'email di promemoria 3 giorni prima di ogni scadenza.");
}

