import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime as dt
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/calendar"]

# Função para criar evento no Google Calendar
def create_calendar_event(summary, location, description, start_datetime, end_datetime, attendees_emails):
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        event = {
            "summary": summary,
            "location": location,
            "description": description,
            "start": {
                "dateTime": start_datetime,
                "timeZone": "UTC"
            },
            "end": {
                "dateTime": end_datetime,
                "timeZone": "UTC"
            },
            "attendees": [{"email": email} for email in attendees_emails]
        }

        event = service.events().insert(calendarId="primary", body=event).execute()
        return event.get('htmlLink')

    except HttpError as error:
        st.error(f"An error occurred: {error}")
        return None

# Mostrar Título e Descrição
st.title("Banco de Validades")
st.markdown("Insira os dados abaixo para armazenar")

# Estabelecendo uma conexão com planilhas do Google
conn = st.connection("gsheets", type=GSheetsConnection)

# Buscar dados de produtos existentes
existing_data = conn.read(worksheet="produtos", usecols=list(range(4)), ttl=6)
existing_data = existing_data.dropna(how="all")

# Onboarding New Vendor Form
with st.form(key="vendor_form"):
    ean_13 = st.number_input(label="EAN-13 do Produto*", step=1, format="%d")
    marca = st.text_input(label="Marca do Produto*")
    data = st.date_input(label="Data de Vencimento*")
    quantidade = st.number_input(label="Quantidade a Vencer*", max_value=24)

    submit_button = st.form_submit_button(label="Submit Vendor Details")

    # If the submit button is pressed
    if submit_button:
        if not ean_13 or not data:
            st.warning("Certifique-se de que todos os campos obrigatórios sejam preenchidos.")
            st.stop()
        elif existing_data["ean_13"].astype(str).str.contains(str(ean_13)).any():
            st.warning("O produto com este EAN-13 já existe.")
            st.stop()
        else:
            if len(str(ean_13)) > 13 or not isinstance(ean_13, int):
                st.warning("O EAN-13 deve ser um número inteiro com até 13 dígitos.")
                st.stop()

            # Adicionar o novo registro ao Google Sheets
            banco_de_datas = pd.DataFrame(
                [
                    {
                        "ean_13": ean_13,
                        "marca": marca,
                        "data": data.strftime("%d-%m-%Y"),
                        "quantidade": quantidade,
                    }
                ]
            )

            updated_df = pd.concat([existing_data, banco_de_datas], ignore_index=True)
            conn.update(worksheet="produtos", data=updated_df)

            # Criar um evento no Google Calendar
            start_datetime = dt.datetime.combine(data, dt.time(9, 0)).isoformat() + "Z"
            end_datetime = dt.datetime.combine(data, dt.time(17, 0)).isoformat() + "Z"
            event_link = create_calendar_event(
                summary=f"Validade do Produto {ean_13}",
                location="Local de Armazenamento",
                description=f"Produto: {marca}\nQuantidade: {quantidade}",
                start_datetime=start_datetime,
                end_datetime=end_datetime,
                attendees_emails=["danilonatan6@gmail.com", "eriknascimento078@gmail.com"]
            )

            if event_link:
                st.success(f"Novo registro enviado com sucesso! [Ver evento no Google Calendar]({event_link})")
            else:
                st.success("Novo registro enviado com sucesso, mas ocorreu um erro ao criar o evento no Google Calendar.")
