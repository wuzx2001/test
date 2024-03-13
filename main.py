import requests
import json
import time
import random 
import os

# Register the azure app first and make sure the app has the following permissions:
# files: Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
# user: User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
# mail: Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
# After registration, you must click on behalf of xxx to grant administrator consent, otherwise outlook api cannot be called


client_id='e90d810c-0af6-400a-bdf2-a6b75e10c8c3'
client_secret='B8K8Q~Fh1ejHCU7hutcx9mJ1QSQUtFoGDUKb1b2.'
refresh_token='0.ASsAq32ZeKne6UCZryyBJXdxVwyBDen2CgpAvfKmt14QyMPCAJg.AgABAAEAAADnfolhJpSnRYB1SVj-Hgd8AgDs_wUA9P-4ibmpdNWRZB-O8dyEVj1VxK3BJ9PSobx9f4ZfNPgbcTmKIJgmLCP7DJZJFJr0hrII2g1Mhpbc_urVhR3CPAXhdFE-XSbR3WhU-winJoV6kuJHXFrpoiA9As9-4hFDFR4NwU4MsLuP2LZL6KMgfHyCQnQ79LIC-u0qefrX6fvmJPT52ZTwIjTq78Vo53mh0TMneVdrMTfkq0TsYjeyyI41vKAg8iqNJSQQF6ShVvzI9HcqwqvsPmXgCeXh4p_RQUg0odHjiJdwI81jhJZ2NFinuPFoGZYFizzEA5-wyDt7oO4AbqzWTKdarBDuniqjKWkCjE0kUnkQMgyDWOWNV6Y1TdszdBI_3QOgijsRAIFvfSjpjyasuor99Upc0aXgHjqhUdxULmsQr-bYtUUHqrN8D1YkFacrLzPA4wlKHVFtobMMqDGNU16zp5gB-EnmGTuqbaSfddimXsCjco71hOF5s4o0aqhHM2KrTTTAbOAsCwp7g4pgLyfTVj2wOHw-bjTzOhNlc7ACrzQLbgn65cKGENREgtiEy_mpDDwOvzVsWjUdmBspfs5Rr0KQ7toT_hDlkBBWwKMdDfkP9whDRaU2k4H6znQmKts8nFR2vWR6bD_OLSJcuPQpbz6aeHkOPRtSTChtw9_4OB9StYiqITaiLg7kMZzvEFvEf6eTYytPog7VzDcgjXCi1O3pbwzpc-7K4hmJoobwDtX7xPaOUukj'



calls = [
    'https://graph.microsoft.com/v1.0/me/drive/root',
    'https://graph.microsoft.com/v1.0/me/drive',
    'https://graph.microsoft.com/v1.0/drive/root',
    'https://graph.microsoft.com/v1.0/users',
    'https://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    'https://graph.microsoft.com/v1.0/me/drive/root/children',
    'https://api.powerbi.com/v1.0/myorg/apps',
    'https://graph.microsoft.com/v1.0/me/mailFolders',
    'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',
    'https://graph.microsoft.com/v1.0/applications?$count=true',
    'https://graph.microsoft.com/v1.0/me/?$select=displayName,skills',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    'https://graph.microsoft.com/beta/me/outlook/masterCategories',
    'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top=1',
    'https://graph.microsoft.com/v1.0/sites/root/lists',
    'https://graph.microsoft.com/v1.0/sites/root',
    'https://graph.microsoft.com/v1.0/sites/root/drives'
]


def get_access_token(refresh_token, client_id, client_secret):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': 'http://localhost:53682/'
    }
    html = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

def main():
    random.shuffle(calls)
    endpoints = calls[random.randint(0,10)::]
    access_token = get_access_token(refresh_token, client_id, client_secret)
    session = requests.Session()
    session.headers.update({
        'Authorization': access_token,
        'Content-Type': 'application/json'
    })
    success = 0
    failed = 0
    for endpoint in endpoints:
        try:
            response = session.get(endpoint)
            if response.status_code == 200:
                success += 1
                print(f'{success+failed}th Call successful')
            else:
                failed +=1
                print(f'{success+failed}th Call failed !!!')
                print("Endpoint :", endpoint)
                print("Response :", response.text)
        except requests.exceptions.RequestException as e:
            failed +=1
            print(f'{success+failed}th request failed !!!')
            print(e)
            pass
    localtime = time.asctime(time.localtime(time.time()))
    print('The end of this run is :', localtime)
    print('Number of successful calls are :', str(success))
    print('Number of failed calls are :', str(failed))

for _ in range(4):
    main()
