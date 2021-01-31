# -*- coding: utf-8 -*-
*** Settings ***
Documentation     Create a Customer Letter and email it
Library           RPA.Robocloud.Secrets
Library           RPA.FileSystem
Library           CreateDocs
Library           RPA.Email.ImapSmtp  smtp_server=smtp.gmail.com      smtp_port=587


*** Variables ***
${GMAIL_ACCOUNT}
${GMAIL_PASSWORD}
${RECIPIENT_ADDRESS}

*** Keywords ***
Set Credential
    ${SECRET}=              Get Secret    credentials
    Set Global Variable     ${GMAIL_ACCOUNT}    ${SECRET}[username]
    Set Global Variable     ${GMAIL_PASSWORD}    ${SECRET}[password]
    Set Global Variable     ${RECIPIENT_ADDRESS}     ${SECRET}[recipient]

*** Variables ***
${WORD_DOCUMENT_PATH}=    ${CURDIR}${/}letter.docx
${PDF_PATH}=              ${CURDIR}${/}letter.pdf
${EMAIL_BODY}=            Please find the document request.
${EMAIL_SUBJECT}=         Bank XYZ - Mortgage Payoff Letter
${ATTACHMENTS}=           ${PDF_PATH}

*** Keywords ***
Send an email
    Set Credential  #https://robocorp.com/docs/development-guide/email/sending-emails-with-gmail-smtp
    Authorize    account=${GMAIL_ACCOUNT}    password=${GMAIL_PASSWORD}
    Send Message  sender=${GMAIL_ACCOUNT}
    ...           recipients=${RECIPIENT_ADDRESS}
    ...           subject=${EMAIL_SUBJECT}
    ...           body=${EMAIL_BODY}
    ...           attachments=${ATTACHMENTS}


*** Keywords ***
Delete files
    Remove file     ${wordDocumentPath}
    Remove file     ${pdfPath}

*** Tasks ***
Orchestrate Task
    #Create a new document and populates the document
    createWordDocument     ${wordDocumentPath}
    createPDF              ${wordDocumentPath}   ${pdfPath}
    Send an email
    #Delete files


