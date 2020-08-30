import email
import imaplib
from dotenv import load_dotenv
from os import environ,path,getcwd

envPath = path.join(getcwd(),'assets','.env_email')
load_dotenv(envPath)

def connectOnEmail():
    emailHost = environ.get("HOST")
    mail = imaplib.IMAP4_SSL(emailHost)
    mail.login(environ.get("EMAIL_USER"),environ.get("EMAIL_PASSWORD"))

    return mail

def disconnectFromEmail(mail):
    mail.close()
    mail.logout()

def getAttacheament(filename,content):

    if filename != None and content != None:

        savePath = path.join(getcwd(),'assets','extraction',filename)

        with open(savePath,'wb') as emailAttacheament:
            emailAttacheament.write(content.get_payload(decode=True))
            emailAttacheament.close()


def searchMessage(mail,subject):
    mail.select()
    typ,data = mail.search(None,'(SUBJECT "{subject}")'.format(subject=subject))

    if typ == 'OK':
        typ,data = mail.fetch(data[0],'(RFC822)')
        emaiBody = data[0][1]
        emailFromBytes = email.message_from_bytes(emaiBody)


        for part in emailFromBytes.walk():
            if part.get_content_maintype() == "multipart":
                continue

            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()

            return filename,part

    return [None,None]

if __name__ == '__main__':

    mail = connectOnEmail()
    filename,part = searchMessage(mail,"Planilha de Repasse")
    getAttacheament(filename,part)
    disconnectFromEmail(mail)


