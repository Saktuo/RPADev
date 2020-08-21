
$path = "C:\Users\saktuo\Desktop\WF_emails.xlsx"

import-module psexcel

$people = new-object System.Collections.ArrayList
$emails = new-object System.Collections.ArrayList

foreach ($person in (Import-XLSX -Path $path -RowStart 1))

{
$people.add($person)
}

foreach ($person in $people){
$emails.add($person.email)
}




$port = 587
$smtp = "smtp.wapice.com"
$mail = new-object System.Net.Mail.MailMessage
$mail.subject = "RPA TEST"
$mail.body = "Hello!

This message was sent by a bot.

The bot is managed by Sakari Tuominen <sakari.tuominen@wapice.com>

Do not reply."
$mail.from = "noreply_rpa@wapice.com"

foreach($email in $emails){
$email
$mail.to.add($email)
}

$mail


