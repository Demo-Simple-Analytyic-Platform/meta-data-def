from modules import secrets as s

def target_db():
    cr = s.credentials("target")
    nm_server   = s.decrypt_data(cr["server"])
    nm_database = s.decrypt_data(cr["database"])
    nm_username = s.decrypt_data(cr["username"])
    nm_password = s.decrypt_data(cr["password"])
    return {
        "server"   : nm_server,
        "database" : nm_database,
        "username" : nm_username,
        "password" : nm_password
    }

def secret_db():
    cr = s.credentials("secrets")
    nm_server   = s.decrypt_data(cr["server"])
    nm_database = s.decrypt_data(cr["database"])
    nm_username = s.decrypt_data(cr["username"])
    nm_password = s.decrypt_data(cr["password"])
    return {
        "server"   : nm_server,
        "database" : nm_database,
        "username" : nm_username,
        "password" : nm_password
    }

def blob_documentation():

    nm_database  = "blob-documentation"
    nm_account   = s.get_secure_information("account",   nm_database)
    nm_secret    = s.get_secure_information("secrret",   nm_database)
    nm_container = s.get_secure_information("container", nm_database)

    return {
        "account"   : s.decrypt_data(nm_account),
        "secret"    : s.decrypt_data(nm_secret),
        "container" : s.decrypt_data(nm_container)
    }
