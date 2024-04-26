# Utilisation du SharePointClient Python pour accéder et télécharger des fichiers et dossiers SharePoint

Dans le monde professionnel moderne, l'intégration de Microsoft SharePoint avec des applications Python permet une gestion automatisée des documents, des dossiers et d'autres ressources de manière programmatique. Ce tutoriel décrit la création et l'utilisation de la classe `SharePointClient` pour interagir avec SharePoint via l'API Microsoft Graph.

Le tutoriel vous guidera à travers l'enregistrement de votre application dans Azure, l'implémentation de la classe Python et la configuration des permissions nécessaires. Enfin, le code source complet sera partagé sur GitHub.

## Prérequis

Avant de commencer, assurez-vous de disposer de :

- Python installé sur votre machine.
- Un accès à un site Microsoft SharePoint.
- La bibliothèque `requests` installée dans Python, disponible via pip (`pip install requests`).

## Étape 1 : Enregistrez votre application

Pour interagir avec SharePoint via l'API Microsoft Graph, vous devez enregistrer votre application dans Azure Active Directory (Azure AD). Cela fournit les `tenant_id`, `client_id` et `client_secret` nécessaires.

### Comment s'enregistrer :

1. **Connectez-vous au portail Azure :** Accédez au [portail Azure](https://portal.azure.com) et connectez-vous.
2. **Accédez à Azure Active Directory :** Sélectionnez Azure Active Directory dans la barre latérale.
3. **Enregistrez une nouvelle application :** Allez dans "Inscriptions d'applications" et cliquez sur "Nouvelle inscription". Fournissez un nom, choisissez les types de comptes et définissez une URI de redirection si nécessaire.
4. **Obtenez les ID et secrets :** Après l'enregistrement, notez l'ID client et l'ID locataire fournis. Créez un nouveau secret client dans "Certificats et secrets".

## Étape 2 : Configurer les permissions

Définissez les permissions correctes dans Azure AD pour permettre à votre application de lire les fichiers et les sites.

### Configuration des permissions :

1. **Permissions de l'API :** Sur la page d'inscription de votre application, cliquez sur "Permissions de l'API".
2. **Ajoutez des permissions :** Sélectionnez "Ajouter une permission", choisissez "Microsoft Graph" puis "Permissions de l'application".
3. **Ajoutez des permissions spécifiques :** Trouvez et ajoutez `Files.Read.All` et `Sites.Read.All` pour activer les capacités de lecture de fichiers et de sites.
4. **Accordez le consentement de l'administrateur :** Pour activer les permissions, cliquez sur "Accorder le consentement administratif pour [Votre Organisation]".

## Étape 3 : Configuration de la classe SharePointClient

Implémentez la classe `SharePointClient` qui inclut l'authentification et les méthodes pour interagir avec les données SharePoint. Ci-dessous, la classe est intégrée dans un script :

```python
import requests
import os

class SharePointClient:
    def __init__(self, tenant_id, client_id, client_secret, resource_url):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.resource_url = resource_url
        self.base_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        self.headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        self.access_token = self.get_access_token()  # Initialize and store the access token upon instantiation

    def get_access_token(self):
        # Body for the access token request
        body = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': self.resource_url + '.default'
        }
        response = requests.post(self.base_url, headers=self.headers, data=body)
        return response.json().get('access_token')  # Extract access token from the response

    def get_site_id(self, site_url):
        # Build URL to request site ID
        full_url = f'https://graph.microsoft.com/v1.0/sites/{site_url}'
        response = requests.get(full_url, headers={'Authorization': f'Bearer {self.access_token}'})
        return response.json().get('id')  # Return the site ID

    def get_drive_id(self, site_id):
        # Retrieve drive IDs and names associated with a site
        drives_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
        response = requests.get(drives_url, headers={'Authorization': f'Bearer {self.access_token}'})
        drives = response.json().get('value', [])
        return [(drive['id'], drive['name']) for drive in drives]

    def get_folder_content(self, site_id, drive_id, folder_path='root'):
        # Get the contents of a folder
        folder_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children'
        response = requests.get(folder_url, headers={'Authorization': f'Bearer {self.access_token}'})
        items_data = response.json()
        rootdir = []
        if 'value' in items_data:
            for item in items_data['value']:
                rootdir.append((item['id'], item['name']))
        return rootdir
    
    # Recursive function to browse folders
    def list_folder_contents(self, site_id, drive_id, folder_id, level=0):
        # Get the contents of a specific folder
        folder_contents_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children'
        contents_headers = {'Authorization': f'Bearer {self.access_token}'}
        contents_response = requests.get(folder_contents_url, headers=contents_headers)
        folder_contents = contents_response.json()

        items_list = []  # List to store information

        if 'value' in folder_contents:
            for item in folder_contents['value']:
                if 'folder' in item:
                    # Add folder to list
                    items_list.append({'name': item['name'], 'type': 'Folder', 'mimeType': None})
                    # Recursive call for subfolders
                    items_list.extend(self.list_folder_contents(site_id, drive_id, item['id'], level + 1))
                elif 'file' in item:
                    # Add file to the list with its mimeType
                    items_list.append({'name': item['name'], 'type': 'File', 'mimeType': item['file']['mimeType']})

        return items_list
    
    def download_file(self, download_url, local_path, file_name):
        headers = {'Authorization': f'Bearer {self.access_token}'}
        response = requests.get(download_url, headers=headers)
        if response.status_code == 200:
            full_path = os.path.join(local_path, file_name)
            with open(full_path, 'wb') as file:
                file.write(response.content)
            print(f"File downloaded: {full_path}")
        else:
            print(f"Failed to download {file_name}: {response.status_code} - {response.reason}")
    
    def download_folder_contents(self, site_id, drive_id, folder_id, local_folder_path, level=0):
        # Recursively download all contents from a folder
        folder_contents_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children'
        contents_headers = {'Authorization': f'Bearer {self.access_token}'}
        contents_response = requests.get(folder_contents_url, headers=contents_headers)
        folder_contents = contents_response.json()

        if 'value' in folder_contents:
            for item in folder_contents['value']:
                if 'folder' in item:
                    new_path = os.path.join(local_folder_path, item['name'])
                    if not os.path.exists(new_path):
                        os.makedirs(new_path)
                    self.download_folder_contents(site_id, drive_id, item['id'], new_path, level + 1)  # Recursive call for subfolders
                elif 'file' in item:
                    file_name = item['name']
                    file_download_url = f"{resource}/v1.0/sites/{site_id}/drives/{drive_id}/items/{item['id']}/content"
                    self.download_file(file_download_url, local_folder_path, file_name)
   
    # Usage example
    tenant_id = 'your-tenant-id'
    client_id = 'your-client-id'
    client_secret = 'your-client-secret'
    site_url = "xxxxx.sharepoint.com:/sites/xxxxxx"  # Replace xxxxx with your site URL
    resource = 'https://graph.microsoft.com/'

    client = SharePointClient(tenant_id, client_id, client_secret, resource)
    site_id = client.get_site_id(site_url)
    print("Site ID:", site_id)

    drive_info = client.get_drive_id(site_id)
    print("Drives available:", drive_info)

    # Example: Access the first drive and list root content
    drive_id = drive_info[0][0]
    folder_content = client.get_folder_content(site_id, drive_id)  
    print("Root Content:", folder_content)
```

## Conclusion

La classe `SharePointClient` offre un moyen simplifié d'interagir avec les ressources SharePoint via Python. Cette solution est idéale pour automatiser les tâches de gestion des documents, améliorant ainsi la productivité au sein de votre organisation. Consultez le code source complet sur [GitHub](https://github.com/ericvaillancourt/Sharepoint-File-Download).

Gardez vos identifiants en sécurité et respectez les meilleures pratiques pour la gestion des informations sensibles. Profitez de l'automatisation avec Python et SharePoint!