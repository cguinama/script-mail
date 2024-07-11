function CreateButton(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Envoyer les mails')
    .addItem('Envoyer les mails', 'approveEmails')
    .addToUi();
}

function approveEmails(){
  var confirme = Browser.msgBox('Confirmation', 'Etes-vous sûr de vouloir envoyer ces mails ? ', browser.Buttons.OK_CANCEL);
  if (confirm == 'ok'){
    sendEmails();
  }
}

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var ccDestinataire = "email@gmail.com";
  var support = "support@gmail.com";

  function formatDate(date) {
    var options = { year: 'numeric', month: 'long'};
    return new Date(date).toLocaleDateString('en-GB', options);
  }
  
  data.forEach(function(row) {
    var id = row[0];
    var date = row[3];
    var email = row[10] + ', ' + row[6] + ', ' + row[4] + ', ' + row[12] + ', ' + support;
    var formattedDate = formatDate(date);

    var sujet = "Information ";
    var message = "Bonjour " + ",\n\n" +
                  "Nous vous informons que l'audit " + id + "ayant eu lieu le  " + formattedDate + " n'a pas été cloturé" + ".\n\n" +
                  "Cordialement,\nVotre équipe";

    MailApp.sendEmail(email, sujet, message, {cc : ccDestinataire});
  });
}
