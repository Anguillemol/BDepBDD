from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
from office365.runtime.client_request import ClientRequest
import io
import pandas as pd

#client_id = "50bbb53b-67ef-488d-9303-d6afcfd77bc8"
#client_secret = "7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc="


site_url = "https://sgzkl-admin.sharepoint.com"

excel_url = "/sites/SiteTest/Documents%20partages/BDD.xlsx"

file_url = site_url + excel_url
#/sites/SiteTest



app_principal = {
    'client_id':'50bbb53b-67ef-488d-9303-d6afcfd77bc8',
    'client_secret':'7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc=',
}

context_auth = AuthenticationContext(url=site_url)
if context_auth.acquire_token_for_app(client_id=app_principal['client_id'], client_secret=app_principal['client_secret']):
    ctx = ClientContext(site_url, context_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Authentification réussie")

response = File.open_binary(ctx, file_url)
df = pd.read_excel(response.content, sheet_name='BDD', engine='openpyxl')

