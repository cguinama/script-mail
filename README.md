Menu personnalisé “Envoyer les mails” :

Cette fonction crée un menu dans la feuille de calcul.
Vous pouvez le trouver sous l’onglet “Extensions” dans Google Sheets.
Cliquez sur “Envoyer les mails” pour exécuter la fonction approveEmails().

Confirmation avant l’envoi des e-mails :
Lorsque vous cliquez sur “Envoyer les mails”, une boîte de dialogue apparaît.
Si vous cliquez sur “OK”, les e-mails seront envoyés.
Si vous cliquez sur “Annuler”, rien ne se produira.

Personnalisation des e-mails :
Les e-mails sont envoyés aux destinataires spécifiés dans la feuille de calcul.
Le sujet de l’e-mail est “Information”.
Le corps de l’e-mail inclut des détails sur l’audit (ID, date) et une note de clôture.

Vous pouvez personnaliser le contenu de l’e-mail en modifiant les variables appropriées dans le code :
sujet : Le sujet de l’e-mail.
message : Le corps de l’e-mail.
ccDestinataire : L’adresse e-mail en copie conforme (CC).

Formatage de la date :
La fonction auxiliaire formatDate(date) formate la date au format lisible par l’homme (par exemple, “janvier 2023”).
Vous pouvez ajuster le format de date en fonction de vos préférences.

Conseils :
Assurez-vous que vos données (ID d’audit, dates, adresses e-mail) sont correctement organisées dans votre feuille de calcul.
Modifiez les variables du code en fonction de vos besoins.
Vérifiez que vous avez les autorisations nécessaires pour envoyer des e-mails via MailApp.sendEmail().
