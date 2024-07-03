function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  function formatDate(date) {
    var options = { year: 'numeric', month: 'long', day: 'numeric' };
    return new Date(date).toLocaleDateString('fr-FR', options);
  }
  
  data.forEach(function(row) {
    var nom = row[0];
    var dateCloture = row[1];
    var email = row[2];
    var ref = row[3];
    var formattedDateCloture = formatDate(dateCloture);

    var sujet = "Information Audit - " + ref;
    var message = "Bonjour " + nom + ",\n\n" +
                  "Nous vous informons que l'audit qui a pour référence" + ref + "ayant eu lieu le  " + formattedDateCloture + " n'a pas été cloturé" + ".\n\n" +
                  "Cordialement,\nVotre équipe";

    MailApp.sendEmail(email, sujet, message);
  });
}
