import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import dropbox
from dropbox.exceptions import AuthError
import os

# Configuración de Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def upload_to_drive(file_path, folder_id=None):
    """
    Sube un archivo a Google Drive y devuelve su enlace público.
    
    :param file_path: Ruta local del archivo a subir
    :param folder_id: ID de la carpeta en Drive (opcional)
    :return: URL pública del archivo o None en caso de error
    """
    try:
        # Autenticación
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        
        # Crear servicio de Drive
        service = build('drive', 'v3', credentials=creds)
        
        # Subir archivo
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [folder_id] if folder_id else []
        }
        media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        # Subir archivo y obtener más campos (incluyendo webViewLink)
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id,webViewLink'
        ).execute()
        
        # Hacer el archivo público
        permission = {
            'type': 'anyone',
            'role': 'reader'
        }
        service.permissions().create(
            fileId=file.get('id'),
            body=permission
        ).execute()
        
        # Obtener el enlace público
        file = service.files().get(
            fileId=file.get('id'),
            fields='webViewLink'
        ).execute()
        
        shared_url = file.get('webViewLink')
        print(f"Archivo subido a Google Drive. Enlace: {shared_url}")
        return shared_url
    
    except Exception as e:
        print(f"Error al subir a Google Drive: {e}")
        return None

def upload_to_dropbox(file_path, dropbox_path):
    
    dbx = dropbox.Dropbox('')
    
    try:
        # Subir el archivo
        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path)
        
        # Obtener el enlace compartido
        link = dbx.sharing_create_shared_link_with_settings(dropbox_path)
        shared_url = link.url  # Este es el enlace público
        
        print(f"Archivo subido a Dropbox. Enlace: {shared_url}")
        return shared_url
    
    except AuthError as e:
        print(f"Error de autenticación: {e}")
        return None
    