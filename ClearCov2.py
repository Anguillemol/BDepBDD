import os
import tempfile
import pandas as pd
import io

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 

username = '50bbb53b-67ef-488d-9303-d6afcfd77bc8'
password = '7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc='

test_team_site_url = "https://sgzkl.sharepoint.com/sites/SiteTest"

ctx = ClientContext(test_team_site_url).with_credentials(ClientCredential(username, password))
file_url = "/sites/SiteTest/Documents%20partages/Test/BDD.xlsx"

#Ca ca marche 
"""
download_path = os.path.join(tempfile.mkdtemp(), os.path.basename(file_url))
with open(download_path, "wb") as local_file:
    file = ctx.web.get_file_by_server_relative_url(file_url).download(local_file).execute_query()

print("[Ok] file has been downloaded into: {0}".format(download_path))
"""

response = File.open_binary(ctx, file_url)

bytes_file_obj = io.BytesIO()
bytes_file_obj.write(response.content)
bytes_file_obj.seek(0)

#Read
df = pd.read_excel(bytes_file_obj, sheet_name='Liste_depots')

print(df)





"""
#import all the libraries
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 
import io
import pandas as pd




url = 'https://sgzkl.sharepoint.com/sites/SiteTest'

username = '50bbb53b-67ef-488d-9303-d6afcfd77bc8'
password = '7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc='



client_credentials = ClientCredential(username, password)
ctx = ClientContext(url).with_credentials(client_credentials)
web = ctx.web
ctx.load(web)
ctx.execute_query()

print(web.properties['Url'])
print(web.properties['WebTemplate'])
print(web.properties['Title'])
"""