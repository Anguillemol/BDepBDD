#import all the libraries
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File 
import io
import pandas as pd


urlSite = "https://sgzkl.sharepoint.com/sites/SiteTest"
cheminAcces = "/Documents%20partages/BDD.xlsx"

username = "lpinbelloc@sgzkl.onmicrosoft.com"
password = "MotDePasseTest31180"

ctx_auth = AuthenticationContext(urlSite)
print("Ici")
if ctx_auth.acquire_token_for_user(username, password):
    print("Puis ici")
    ctx = ClientContext(urlSite, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Autehtification réussie")

print("Chelou mec")
reponse = File.open_binary(ctx, urlSite)