from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import smtplib
from email.message import EmailMessage
import json
import requests
import pyodbc
from datetime import datetime, timedelta, timezone
import os, base64
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.backends import default_backend

def encrypt(plaintext: str, key_b64: str) -> str:
    """
    Krypter plaintext med AES-CBC + PKCS7.
    key_b64: Base64-encoded nøgle (skal give 16, 24 eller 32 bytes efter decode).
    Returnerer: Base64-encoded IV + ciphertext.
    """
    key = base64.b64decode(key_b64)
    iv = os.urandom(16)

    padder = padding.PKCS7(128).padder()
    padded_data = padder.update(plaintext.encode()) + padder.finalize()

    cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
    encryptor = cipher.encryptor()
    ciphertext = encryptor.update(padded_data) + encryptor.finalize()

    return base64.b64encode(iv + ciphertext).decode()

def insert_new_case(cur, data, IndsenderNavn, IndsenderID, IndsenderMail, AnmodningsID, Beskrivelse):
    # 1) cases
    cur.execute("""
        INSERT INTO dbo.cases (citizen_name, citizen_id, citizen_email, status, PersonaleSagsTitel, Beskrivelse, AktID)
        OUTPUT INSERTED.id
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (IndsenderNavn, IndsenderID, IndsenderMail, "Ny", f'Anmodning {AnmodningsID}' , Beskrivelse, AnmodningsID))
    case_id = cur.fetchone()[0]

    # 2) case_journal_items (received)
    cur.execute("""
        INSERT INTO dbo.case_journal_items (case_id, item_type, payload, journal_status)
        VALUES (?, ?, ?, DEFAULT)
    """, (case_id, "received", json.dumps(data, ensure_ascii=False)))

    # 3) caselogs  — INKLUDÉR ET TZ-AWARE TIMESTAMP (UTC)
    utc_now = datetime.now(timezone.utc)
    cur.execute("""
        INSERT INTO dbo.caselogs ([case_id], [message], [field], [action], [user], [timestamp])
        VALUES (?, ?, ?, ?, ?, ?)
    """, (case_id, "Sag modtaget via formular", "status", "modtaget", "System", utc_now))

    return AnmodningsID

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    orchestrator_connection.log_info('Started proces EmailNyPersonaleAktindsigt')

    specific_content = json.loads(queue_element.data)
    AnmodningsID = specific_content.get('application_uuid')

    os2forms_user = orchestrator_connection.get_credential('OS2FormsAPI')
    os2formsURL = os2forms_user.username
    os2formsApiKey = os2forms_user.password
    encryptionkey = orchestrator_connection.get_credential('PersonalesagsEncryptionKey').password

    url = f"{os2formsURL}laura_salmonsen_aktindsigt_test/submission/{AnmodningsID}"

    headers = {
    'api-key': f'{os2formsApiKey}'
    }

    response = requests.get( url, headers=headers)
    response.raise_for_status()
    entity = response.json()['entity']
    data = response.json()['data']


    IndsenderNavn = encrypt(data.get('navn_paa_ansoeger'), encryptionkey)
    IndsenderMail = data.get('email')
    IndsenderID = encrypt(data.get('cpr_nummer_paa_ansoeger'), encryptionkey)
    ModtagerMail = orchestrator_connection.get_constant('balas').value #Ændr til rigtig modtagermail fra HR
    AktID = entity.get('sid')[0].get('value')
    ModtagerTekst = data.get('her_kan_du_konkretisere_din_anmodning', "")
    dato_string = entity.get('completed')[0].get('value')
    IndsendelsesDato = datetime.fromisoformat(dato_string).strftime("%d-%m-%Y %H:%M")

    if any(x is None for x in [IndsenderNavn, IndsenderMail, IndsenderID, AktID, IndsendelsesDato]):
        orchestrator_connection.log_info('Missing information in application')
        raise Exception

    #----------------- Here the case details are sent to the database
    sql_server = orchestrator_connection.get_constant("SqlServer").value  
    conn_string = f"DRIVER={{SQL Server}};SERVER={sql_server};DATABASE=AKTINDSIGTERPERSONALEMAPPER;Trusted_Connection=yes;"
    conn = pyodbc.connect(conn_string)
    conn.autocommit = False
    cur = conn.cursor()
    case_id = insert_new_case(cur, data, IndsenderNavn, IndsenderID, IndsenderMail, AktID, ModtagerTekst)
    conn.commit()
    orchestrator_connection.log_info(f"Oprettet sag id={AktID}")

    # ---------------- Here mail to applicant and sagsbehandler is sent
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "PersonaleAktindsigtssag@aarhus.dk"
    subject_sagsbehandler = "Ny anmodning om aktindsigt i personalesag"

    html = f"""
    <html>
    <body>
        <p>Der er den {IndsendelsesDato} indsendt en ny anmodning om aktindsigt i en personalesag med AktID {AktID}. </p>
        <p>Du kan se sagen på linket herunder: </p>
        <p> LINK til sagen skal indsættes </p> 
    </body>
    </html>
    """
    # Create the email message
    UdviklerMail = orchestrator_connection.get_constant('balas').value

    msg = EmailMessage()
    msg['To'] = ModtagerMail
    msg['From'] = SCREENSHOT_SENDER
    msg['Subject'] = subject_sagsbehandler
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(html, subtype='html')
    msg['Reply-To'] = UdviklerMail
    msg['Bcc'] = UdviklerMail

    # SMTP Configuration (from your provided details)
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "PersonaleAktindsigtssag@aarhus.dk"
    subject_anmoder = "Kvittering for modtagelse af anmodning om aktindsigt"

    html_anmoder = f"""
    <html>
    <body>
        <p>Kære {IndsenderNavn}, </p>
        <p>Vi har den {IndsendelsesDato} modtaget din anmodning om aktindsigt i din personalemappe, og har givet anmodningen ID {AktID}. </p>
        <p>En medarbejder vil gå i gang med at se på din anmodning </p>
    </body>
    </html>
    """
    msg_anmoder = EmailMessage()
    msg_anmoder['To'] = IndsenderMail
    msg_anmoder['From'] = SCREENSHOT_SENDER
    msg_anmoder['Subject'] = subject_anmoder
    msg_anmoder.set_content("Please enable HTML to view this message.")
    msg_anmoder.add_alternative(html_anmoder, subtype='html')
    msg_anmoder['Reply-To'] = UdviklerMail
    msg_anmoder['Bcc'] = UdviklerMail

    # Send the email using SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg)
            smtp.send_message(msg_anmoder)
    except Exception as e:
        orchestrator_connection.log_info(f"Failed to send success email: {e}")
