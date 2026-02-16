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
def tjek_email(cur, citizen_id, citizen_email):
    
    birthday = citizen_id[:6]

    cur.execute("""
        SELECT Email
        FROM FDW.pdb.PersonLight_udvidet
        WHERE Fødselsdag = ?
    """, birthday)

    rows = cur.fetchall()

    # Lav liste med kun emails
    emails = [row[0] for row in rows]
    if citizen_email in emails:
        return True
    else:
        return False

def get_next_aktid(cur) -> str:
    year = datetime.now().year

    # Lås rækken for det pågældende år, så kun én transaktion ad gangen kan opdatere
    cur.execute("""
        SELECT current_value
        FROM dbo.case_id_counter WITH (UPDLOCK, HOLDLOCK)
        WHERE year = ?
    """, (year,))
    row = cur.fetchone()

    if row is None:
        current_value = 1
        cur.execute("""
            INSERT INTO dbo.case_id_counter (year, current_value)
            VALUES (?, ?)
        """, (year, current_value))
    else:
        current_value = row[0] + 1
        cur.execute("""
            UPDATE dbo.case_id_counter
            SET current_value = ?
            WHERE year = ?
        """, (current_value, year))

    # Format: ÅÅÅÅ-XXXX, justér som du vil
    return f"{year}-{current_value:04d}"

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
        INSERT INTO dbo.cases (
            aktid, citizen_name, citizen_id, citizen_email, status, PersonaleSagsTitel, Beskrivelse
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        AnmodningsID,
        IndsenderNavn,
        IndsenderID,
        IndsenderMail,
        "Ny",
        AnmodningsID,
        Beskrivelse
    ))

    # 2) case_journal_items (received)
    cur.execute("""
        INSERT INTO dbo.case_journal_items (case_aktid, item_type, payload, journal_status)
        VALUES (?, ?, ?, DEFAULT)
    """, (
        AnmodningsID,
        "received",
        json.dumps(data, ensure_ascii=False)
    ))

    # 3) caselogs — inkl. UTC timestamp
    utc_now = datetime.now(timezone.utc)
    cur.execute("""
        INSERT INTO dbo.caselogs (case_aktid, message, field, action, [user], [timestamp])
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        AnmodningsID,
        "Sag modtaget via formular",
        "status",
        "modtaget",
        "System",
        utc_now
    ))

    return AnmodningsID

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    orchestrator_connection.log_info('Started proces EmailNyPersonaleAktindsigt')

    specific_content = json.loads(queue_element.data)
    AnmodningsID = specific_content.get('application_uuid')

    os2forms_user = orchestrator_connection.get_credential('OS2FormsAPI')
    os2formsURL = os2forms_user.username
    os2formsApiKey = os2forms_user.password
    encryptionkey = os.getenv('PersonaleIndsigtEncryptionKey')
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "PersonaleAktindsigtssag@aarhus.dk"

    url = f"{os2formsURL}laura_salmonsen_aktindsigt_test/submission/{AnmodningsID}"

    headers = {
    'api-key': f'{os2formsApiKey}'
    }

    response = requests.get( url, headers=headers)
    response.raise_for_status()
    entity = response.json()['entity']
    data = response.json()['data']

    IndsenderNavn = data.get('navn_paa_ansoeger')
    IndsenderMail = data.get('email')
    IndsenderIDraw = data.get('cpr_nummer_paa_ansoeger')
    IndsenderID = encrypt(data.get('cpr_nummer_paa_ansoeger'), encryptionkey)
    ModtagerMail = orchestrator_connection.get_constant('balas').value #Ændr til rigtig modtagermail fra HR
    sid = entity.get('sid')[0].get('value')
    ModtagerTekst = encrypt(data.get('her_kan_du_konkretisere_din_anmodning', ""), encryptionkey)
    dato_string = entity.get('completed')[0].get('value')
    IndsendelsesDato = datetime.fromisoformat(dato_string).strftime("%d-%m-%Y %H:%M")

    data_for_journal = data.copy()
    data_for_journal["cpr_nummer_paa_ansoeger"] = IndsenderID


    if any(x is None for x in [IndsenderNavn, IndsenderMail, IndsenderID, sid, IndsendelsesDato]):
        orchestrator_connection.log_info('Missing information in application')
        raise Exception

    #First we check if the email is right - if not, citizen receives email
    sql_server_f = orchestrator_connection.get_constant("sqlserverf").value  
    conn_string_f = f"DRIVER={{SQL Server}};SERVER={sql_server_f};DATABASE=FDW;Trusted_Connection=yes;"
    conn_f = pyodbc.connect(conn_string_f)
    conn_f.autocommit = False
    cur_f = conn_f.cursor()
    if not tjek_email(cur_f, citizen_id= IndsenderIDraw, citizen_email= IndsenderMail):
      
        orchestrator_connection.log_info('Mail does not correspond to birthday. Applicant emailed')
        # ---------------- Here mail to applicant and sagsbehandler is sent
        html_failed = f"""
        <html>
        <body>
            <p>Den angivne mailadresse på din aktindsigtsanmodning matcher ikke til cpr-nummer. Tjek at du har brugt din egen aarhus-mail. </p>
            <p>Kontakt HR hvis du har brug for hjælp.</p>
        </body>
        </html>
        """
        # Create the email message
        UdviklerMail = orchestrator_connection.get_constant('balas').value

        msg_failed = EmailMessage()
        msg_failed['To'] = IndsenderMail
        msg_failed['From'] = SCREENSHOT_SENDER
        msg_failed['Subject'] = "Forkert mailadresse angivet"
        msg_failed.set_content("Please enable HTML to view this message.")
        msg_failed.add_alternative(html_failed, subtype='html')

        # Send the email using SMTP
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.send_message(msg_failed)
        except Exception as e:
            orchestrator_connection.log_info(f"Failed to send error email: {e}")

    else:
        #----------------- Here the case details are sent to the database
        sql_server = orchestrator_connection.get_constant("SqlServer").value  
        conn_string = f"DRIVER={{SQL Server}};SERVER={sql_server};DATABASE=AKTINDSIGTERPERSONALEMAPPER;Trusted_Connection=yes;"
        conn = pyodbc.connect(conn_string)
        conn.autocommit = False
        cur = conn.cursor()
        aktid = get_next_aktid(cur)
        aktid = insert_new_case(cur, data_for_journal, IndsenderNavn, IndsenderID, IndsenderMail, aktid, ModtagerTekst)
        conn.commit()
        orchestrator_connection.log_info(f"Oprettet sag med aktid={aktid}")

        # ---------------- Here mail to applicant and sagsbehandler is sent
        subject_sagsbehandler = "Ny anmodning om aktindsigt i personalesag"

        html = f"""
        <html>
        <body>
            <p>Der er den {IndsendelsesDato} indsendt en ny anmodning om aktindsigt i en personalesag med AktID {aktid}. </p>
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

        subject_anmoder = "Kvittering for modtagelse af anmodning om aktindsigt"

        html_anmoder = f"""
        <html>
        <body>
            <p>Kære {IndsenderNavn}, </p>
            <p>Vi har den {IndsendelsesDato} modtaget din anmodning om aktindsigt i din personalemappe, og har givet anmodningen ID {aktid}. </p>
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
