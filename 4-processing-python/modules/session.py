# Import Credentials
from modules import credentials as cr

# Import for Spark for session
from pyspark.sql import SparkSession

def getSparkSession(nm_application):
    
    # Initialize Spark session
    spark = SparkSession.builder\
           .appName(nm_application)\
           .config("spark.pyspark.python",        cr.ds_path_to_python) \
           .config("spark.pyspark.driver.python", cr.ds_path_to_python) \
           .config("spark.driver.extraClassPath", cr.ds_path_to_jdbc) \
           .getOrCreate()
           #.config("spark.sql.execution.arrow.pyspark.enabled", "true") \
           
    # Return the Spark session
    return spark

def jdbc_url(
    
    # Input Parameters
    dc_credentials_db # database credentials

):
    ds_jdbc_url  = f"jdbc:sqlserver://{dc_credentials_db['server']};"
    ds_jdbc_url += f"databaseName={dc_credentials_db['database']};"
    ds_jdbc_url += f"encrypt=false;trustServerCertificate=false"

    return ds_jdbc_url

def truncate_table(
    
    # Input Parameters
    dc_credentials_db, # database credentials
    nm_target_schema,  # target schema name
    nm_target_table,   # target table name    

):
    
    # Set Credentials
    ds_jdbc_url = jdbc_url(dc_credentials_db)
    nm_username = dc_credentials_db['username']
    cd_password = dc_credentials_db['password']

    # Start Spark Session for Truncate Table
    spark_truncate = getSparkSession("TruncateTableFunction")
    tx_truncate_table = f"TRUNCATE TABLE {nm_target_schema}.{nm_target_table}"
    driver_manager = spark_truncate._sc._gateway.jvm.java.sql.DriverManager
    connection = driver_manager.getConnection(ds_jdbc_url, nm_username, cd_password)
    connection.prepareCall(tx_truncate_table).execute()
    connection.close()

    # Stop Spark Session
    spark_truncate.stop()
    
# Test getSparkSession function
#spark = getSparkSession('Test')

