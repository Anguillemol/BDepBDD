from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
from office365.runtime.client_request import ClientRequest

import io
import os
import tempfile
import pandas as pd

#Permission de lecture ecriture sur le site BricoDepot
test_team_site_url = "https://sgzkl.sharepoint.com/sites/BricoDepot"


client_id = "acae250d-01e9-4f32-9d65-e06fa388ff60"
client_secret = "8FG7d+Es/DYXCJWN8spbNV6qyU5TQqUsoKmg5HLsHw4="
title = "MyApp2"

app_principal = {
    'client_id': client_id,
    'client_secret':client_secret,
}
"""
app_principal = {
    'client_id':'50bbb53b-67ef-488d-9303-d6afcfd77bc8',
    'client_secret':'7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc=',
}
"""
context_auth = AuthenticationContext(url=test_team_site_url)
if context_auth.acquire_token_for_app(client_id=app_principal['client_id'], client_secret=app_principal['client_secret']):
    ctx = ClientContext(test_team_site_url, context_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Authentification réussie")

file_url = "/sites/BricoDepot/Shared%20Documents/Donnees/BDD.xlsx"
download_path = os.path.join(tempfile.mkdtemp(), os.path.basename(file_url))
with open(download_path, "wb") as local_file:
    file = ctx.web.get_file_by_server_relative_url(file_url).download(local_file).execute_query()
    #file = ctx.web.get_file_by_server_relative_url(file_url).download(local_file).execute_query()
print("[Ok] file has been downloaded into: {0}".format(download_path))

path = download_path

