def lancer_traitement(envoi_actif=True, callback_log=None):
    from docxtpl import DocxTemplate
    from docx2pdf import convert
    from datetime import datetime
    import pandas as pd
    import os
    import smtplib
    import ssl
    from email.message import EmailMessage
    import time
    import unicodedata
    import re
    import locale

    # === CHEMINS ===
    DOSSIER_BASE = os.path.dirname(os.path.abspath(__file__))
    FICHIER_EXCEL = os.path.join(DOSSIER_BASE, 'adherents.xlsx')
    MODELE_WORD = os.path.join(DOSSIER_BASE, 'modele.docx')
    DOSSIER_RESULTATS = os.path.join(DOSSIER_BASE, 'resultats')
    CHEMIN_LOG = os.path.join(DOSSIER_RESULTATS, 'logs.txt')
    os.makedirs(DOSSIER_RESULTATS, exist_ok=True)

    # === CONFIGURATION EMAIL ===
    EMAIL_EXPEDITEUR = "michelpiallier@gmail.com"
    MOT_DE_PASSE = "ohlvtusqvnxwfsap"
    SMTP_SERVEUR = "smtp.gmail.com"
    SMTP_PORT = 465
    
    # === FONCTIONS ===
    def ecrire_log(message):
        horodatage = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        ligne = f"{horodatage} {message}"
        with open(CHEMIN_LOG, "a", encoding="utf-8") as f:
            f.write(ligne + "\n")
        if callback_log:
            callback_log(ligne)

    def nettoyer_nom_fichier(nom):
        nom = unicodedata.normalize('NFKD', nom).encode('ASCII', 'ignore').decode('utf-8')
        nom = nom.replace(' ', '-')
        nom = re.sub(r'[^a-zA-Z0-9\\-]', '', nom)
        return nom.lower()

    def envoyer_email(destinataire, civilite, prenom, nom, cotisation, pdf_path):
        msg = EmailMessage()
        msg['From'] = EMAIL_EXPEDITEUR
        msg['To'] = destinataire
        msg['Subject'] = "Votre reçu fiscal"
        corps = (
            f"Bonjour {civilite} {prenom} {nom},\n\n"
            "Chère adhérente, cher adhérent,\n\n"
            "Je vous prie de trouver, en annexe, le reçu fiscal de l'année 2024.\n"
            "Je vous en souhaite bonne réception.\n\n"
            "Cordialement,"
)
        msg.set_content(corps)
        with open(pdf_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVEUR, SMTP_PORT, context=context) as server:
            server.login(EMAIL_EXPEDITEUR, MOT_DE_PASSE)
            server.send_message(msg)

    def email_valide(email):
        email = email.strip()
        pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
        return re.match(pattern, email) is not None

    # === LECTURE DU FICHIER EXCEL ===
    df = pd.read_excel(FICHIER_EXCEL, dtype=str).fillna("")
    mappage_colonnes = {
        "Civilité": "civilite",
        "NOM": "nom",
        "PRENOM ADH": "prenom",
        "N° ORDRE": "num_ordre",
        "Somme": "cotisation_1",
        "Somme en lettres": "cotisation_2",
        "E-mail": "email",
        "ADRESSE 1": "adresse_1",
        "ADRESSE 2": "adresse_2",
        "C.P.": "code_postal",
        "VILLE": "ville",
        "DATE": "date_paiement"
    }
    df.rename(columns=mappage_colonnes, inplace=True)

    df["email"] = df["email"].astype(str).str.replace('\\n', '').str.replace('\\r', '').str.replace('\\t', '').str.strip()

    emails_invalides = []

    def filtrer_email(row):
        email = str(row.get("email", "")).replace('\\n', '').replace('\\r', '').replace('\\t', '').strip()
        if not email_valide(email):
            emails_invalides.append(f"{row.get('prenom', '')} {row.get('nom', '')} → {email}")
            return False
        return True

    df = df[df.apply(filtrer_email, axis=1)]

    if emails_invalides:
        for ligne in emails_invalides:
            ecrire_log(f"⛔ Email invalide ignoré : {ligne}")
        ecrire_log(f"{len(emails_invalides)} email(s) invalides ignorés.")

    df["code_postal"] = df["code_postal"].apply(lambda x: str(x).split('.')[0].zfill(5))
    df["num_ordre"] = df["num_ordre"].apply(lambda x: f"{int(float(x)):03d}" if str(x).strip() != "" else "000")
    df["cotisation_1"] = df["cotisation_1"].apply(lambda x: f"{int(float(x)):,}".replace(",", " ") + " €" if str(x).strip() != "" else "0 €")

    for _, row in df.iterrows():
        donnees = row.to_dict()
        donnees["date_envoi"] = datetime.now().strftime("%d/%m/%Y")
        locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')

        try:
            date_obj = pd.to_datetime(donnees.get("date_paiement", ""), errors='coerce')
            if pd.notna(date_obj):
                donnees["date_paiement"] = date_obj.strftime("%d/%m/%Y")
            else:
                donnees["date_paiement"] = ""
        except:
            donnees["date_paiement"] = ""

        nom = nettoyer_nom_fichier(donnees["nom"])
        prenom = nettoyer_nom_fichier(donnees["prenom"])
        fichier_base = f"recu-fiscal-{prenom}-{nom}"
        chemin_docx = os.path.join(DOSSIER_RESULTATS, fichier_base + ".docx")
        chemin_pdf = os.path.join(DOSSIER_RESULTATS, fichier_base + ".pdf")

        try:
            doc = DocxTemplate(MODELE_WORD)
            doc.render(donnees)
            doc.save(chemin_docx)
            convert(chemin_docx, chemin_pdf)

            if envoi_actif:
                envoyer_email(donnees["email"], donnees["civilite"], donnees["prenom"], donnees["nom"], donnees["cotisation_1"], chemin_pdf)
                ecrire_log(f"✅ Email envoyé à {donnees['email']} ({donnees['prenom']} {donnees['nom']})")
            else:
                ecrire_log(f"🧪 TEST : fichier généré pour {donnees['email']} ({donnees['prenom']} {donnees['nom']})")
                #ecrire_log(f"Date pour {donnees['prenom']} {donnees['nom']} : {donnees['date_paiement']}")

            time.sleep(1)
        except Exception as e:
            ecrire_log(f"❌ Erreur pour {donnees['email']} ({donnees['prenom']} {donnees['nom']}) : {str(e)}")
