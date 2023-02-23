#import all the libraries
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
from office365.runtime.auth.client_credential import ClientCredential
import io
import pandas as pd

cheminAcces = "/Documents%20partages/BDD.xlsx"

url = "https://sgzkl.sharepoint.com/sites/SiteTest"



client_id = '50bbb53b-67ef-488d-9303-d6afcfd77bc8'
client_secret = '7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc='








"""
https://sgzkl.sharepoint.com/_layouts/15/appregnew.aspx
appregnew
appinv.aspx
secretID = "50bbb53b-67ef-488d-9303-d6afcfd77bc8"
secretPassWord = "7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc="


#Name is PythonConsole
"""
username = "lpinbelloc@sgzkl.onmicrosoft.com"
password = "MotDePasseTest31180"

ctx_auth = AuthenticationContext(url)
print("Ici")
if ctx_auth.acquire_token_for_user(username, password):
    print("Puis ici")
    ctx = ClientContext(url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Autehtification réussie")

print("Chelou mec")
reponse = File.open_binary(ctx, url)


