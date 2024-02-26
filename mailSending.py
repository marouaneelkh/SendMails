import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart



def sendMails(destinataire, sujet, corps):
    # Params SMTP (Simple mail transfer protocol)
    serverSmtp = "smtp.gmail.com"
    serverPort = 587


    # UserId
    email = "votremail@gmai/com"
    
    # Si vous avez activé l'authentification à 2 facteurs 
    # Créer mdp d'applications sur les params de votre mail
    # Sinon Entrez directement votre mdp réel
    mdp = "votreMdp"

    
    message = MIMEMultipart()


    message['From'] = email
    message['To'] = destinataire 
    message['Subject'] = sujet



    message.attach(MIMEText(corps, 'plain'))

     # Établir la connexion avec le serveur SMTP
    server = smtplib.SMTP(serverSmtp, serverPort)
    server.starttls()


    server.login(email, mdp)

    # Envoyer l'e-mail
    server.sendmail(email, destinataire, message.as_string())

    # Fermer la connexion
    server.quit()





dest = "destinataire@gmail.com"

suj = "Sujet : Mail"

corps = "corps de votre Email"



# Charger le fichier Excel
fichier_excel = 'testMail.xlsx'
classeur = openpyxl.load_workbook(fichier_excel)
feuille = classeur.worksheets[0]


# Parcourir les lignes du fichier Excel
for ligne in feuille.iter_rows(min_row=2,values_only=True):
    nom_entreprise, email = ligne[0], ligne[1]
    print(email)
    
    sujet_email = 'Test_Envoi_Mails'
    
    corps_email = f"Bonjour {nom_entreprise},\n\nCeci est le corps de votre email."

    # Appeler la fonction pour envoyer l'e-mail
    sendMails(email, sujet_email, corps_email)

print("E-mails envoyés avec succès.")





    
    