# Python snowflake connector test file
import snowflake.connector
import logging
import requests
import json
# Azure user
# USER = "DEV_USER@prashantgsharmakipi.onmicrosoft.com"
# PASSWORD = "Kipi.bi@2022""
USER = "rohit.pawar@biourja.com"
PASSWORD = "Biourja@123456#"
# Snowflake options
ACCOUNT = "os54042.east-us-2.azure"
WAREHOUSE = "BUPOWER_INDIA_WH"
DATABASE = "POWERDB_DEV"
SCHEMA = "KIPI_DEV"
ROLE = "OWNER_POWERDB_DEV" 
#ROLE = "READER_GLOBAL" 
# Azure AD options
AUTH_CLIENT_ID = "d75b2b11-f111-4117-a1c4-e56493cc8034"
AUTH_CLIENT_SECRET = "wtL8Q~v25M8To5l7IoFWPnjL66h5uLCWqjFilb-t"
AUTH_GRANT_TYPE = "password"
SCOPE_URL = "api://snowflake.biourja.com/8a32c9df-d1d4-42c1-91af-671eb66c7fc5/session:scope:OWNER_POWERDB_DEV"
TOKEN_URL = "https://login.microsoftonline.com/8ded9ee3-9568-4108-94a1-774f93131d6f/oauth2/v2.0/token"
PAYLOAD = "client_id={clientId}&" \
          "client_secret={clientSecret}&" \
          "username={userName}&" \
          "password={password}&" \
          "grant_type={grantType}&" \
          "scope={scopeUrl}".format(clientId=AUTH_CLIENT_ID, clientSecret=AUTH_CLIENT_SECRET, userName=USER,
                                    password=PASSWORD, grantType=AUTH_GRANT_TYPE, scopeUrl=SCOPE_URL)
logging.basicConfig(
            filename="log.log",
            level=logging.DEBUG)
print("Getting JWT token")
response = requests.post(TOKEN_URL, data=PAYLOAD)
json_data = json.loads(response.text)
print('Json Data: ', json_data)
TOKEN = json_data['access_token']
print("Token obtained", TOKEN)
# Snowflake connection
print("connecting to Snowflake")
conn = snowflake.connector.connect(
                user=USER,
                account=ACCOUNT,
                role=ROLE,
                authenticator="oauth",
                token=TOKEN,
                warehouse=WAREHOUSE,
                database=DATABASE,
                schema=SCHEMA
                )
cur = conn.cursor()
print("connected to snowflake")
try:
    print("running command")
    cur.execute("select current_version();")
    ret = cur.fetchone()[0]
    print('Current Version: ', ret)
except snowflake.connector.errors.ProgrammingError as e:
    print(e)
finally:
    print("closing connection to snowflake")
    conn.close()