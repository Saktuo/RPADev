*** Settings ***
Documentation   Sending email via Gmail
Library         ExampleLibrary
Library         RPA.Email.ImapSmtp    smtp_server=smtp.gmail.com    smtp_port=587
Variables       variables.py

# +
*** Variables ***

${RECIPIENT}      sakari.tuominen@wapice.com
