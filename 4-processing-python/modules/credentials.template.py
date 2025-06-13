# Description: This file contains the credentials for the database connections. 
# to make the code work, a copy must be made of this file and named credentials.py.
# The credentials must be filled in with the correct values.
#
# The credentials for the "Secret"-database.
secret_db = {
    "server"   : "server_name,port",
    "database" : "secrets",
    "username" : r"username",
    "password" : r"password"
}

# The credentials for the "target"-database.
target_db = {
    "server"   : "server_name,port",
    "database" : "secrets",
    "username" : r"username",
    "password" : r"password"
}

# The credentials for the "documentation"-blob storage with is setup as a static-webserver.
blob_documentation = {
    "account"   : "your-accoutn-name",
    "secret"    : "the-secret-name-of-the-accesskey",
    "container" : "your-container-name"
}

# path to the executable of Python
ds_path_to_python = r'path_to_and_including_python.exe'

# path to the JDBC SQL Server Drivers (Jars)
ds_path_to_jdbc   = r'path_to_jdbc_sql_server_jar\sqljdbc_12.8\enu\jars\mssql-jdbc-12.8.1.jre8.jar'

