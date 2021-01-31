# Generate-PDF-Letter
Generate a Mortgage Payoff Letter - Word, PDF and email the pdf as an attachment.

<b>Dependencies</b>: Robocorp Lab, local vault setup, docx2pdf(pip library), python-docx(conda library). 
The dependencies are already packaged in this solution<br>
FYI, this has only been tested on MacOS and may or may not work for other operating systems.

<h2> How it works? </h2>

<b>Task.robots</b>: orchestrates the process of creating a word document, creating a pdf, emailing the pdf and deleting files 

<b>Libraries</b>: contains the CreateDocs.py with methods to create word doc and pdf

<b>devdata</b>: contains the env.json (for environment variables such as the gmail account credentials).
The env.json -> "RPA_SECRET_FILE" needs to be configured to link to your local json vault.
A sample vault.json
```
{
  "credentials": {
    "username": "any email account",
    "password": "your email password",
    "recipient": "recipient email address"
  }
}
```

Configure local vault - https://robocorp.com/docs/development-guide/variables-and-secrets/vault <br>
Send smtp email - https://robocorp.com/docs/development-guide/email/sending-emails-with-gmail-smtp
