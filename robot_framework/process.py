from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import os
import smtplib
from email.message import EmailMessage
import json
import requests
import pyodbc

def insert_new_case(cur, data, IndsenderNavn, IndsenderID, IndsenderMail):
    # 1) cases
    cur.execute("""
        INSERT INTO dbo.cases (citizen_name, citizen_id, citizen_email, status, PersonaleSagsTitel)
        OUTPUT INSERTED.id
        VALUES (?, ?, ?, ?, ?)
    """, (IndsenderNavn, IndsenderID, IndsenderMail, "Modtaget", "Aktindsigt i personalemappe"))
    case_id = cur.fetchone()[0]

    # 2) case_journal_items (received)
    cur.execute("""
        INSERT INTO dbo.case_journal_items (case_id, item_type, payload, journal_status)
        VALUES (?, ?, ?, DEFAULT)
    """, (case_id, "received", json.dumps(data, ensure_ascii=False)))

    # 3) caselogs
    cur.execute("""
        INSERT INTO dbo.caselogs (case_id, message, field, action, user)
        VALUES (?, ?, ?, ?, ?)
    """, (case_id, "Sag modtaget via formular", "status", "modtaget", "System"))

    return case_id

def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    orchestrator_connection.log_info('Started proces EmailNyPersonaleAktindsigt')
    specific_content = json.loads(queue_element.data)
    AnmodningsID = specific_content.get('application_uuid')

    os2forms_user = orchestrator_connection.get_credential('OS2FormsAPI')
    os2formsURL = os2forms_user.username
    os2formsApiKey = os2forms_user.password
    

    url = f"{os2formsURL}laura_salmonsen_aktindsigt_test/submission/{AnmodningsID}"

    headers = {
    'api-key': f'{os2formsApiKey}'
    }

    response = requests.get( url, headers=headers)
    response.raise_for_status()
    data = response.json()['data']
    

    IndsenderNavn = data.get('citizen_name')
    IndsenderMail = data.get('citizen_mail')
    IndsenderID = data.get('citizen_id')
    ModtagerMail = orchestrator_connection.get_constant('balas').value #Ændr til rigtig modtagermail fra HR
    AktID = specific_content.get('application_id')
    IndsendelsesDato = specific_content.get('application_date')

    if any(x is None for x in [IndsenderNavn, IndsenderMail, IndsenderID, AktID, IndsendelsesDato]):
        orchestrator_connection.log_info('Missing information in application')
        raise Exception
    
    #----------------- Here the case details are sent to the database
    sql_server = orchestrator_connection.get_constant("SqlServer").value  # fx "db01" el. "db01,1433"
    conn_string = f"DRIVER={{SQL Server}};SERVER={sql_server};DATABASE=AKTINDSIGTIPERSONALEMAPPER;Trusted_Connection=yes;"
    conn = pyodbc.connect(conn_string)
    conn.autocommit = False
    cur = conn.cursor()
    case_id = insert_new_case(cur, data, IndsenderNavn, IndsenderID, IndsenderMail)
    conn.commit()
    orchestrator_connection.log_info(f"Oprettet sag id={case_id}")

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
    msg['From'] = SCREENSHOT_SENDER
    msg['Subject'] = subject_anmoder
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(html_anmoder, subtype='html')
    msg['Reply-To'] = UdviklerMail
    msg['Bcc'] = UdviklerMail

    # Send the email using SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg)
            smtp.send_message(msg_anmoder)
    except Exception as e:
        orchestrator_connection.log_info(f"Failed to send success email: {e}")
