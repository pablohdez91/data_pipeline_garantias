import polars as pl
import pyodbc

def read_access(file, table):
    try:
        conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + file
        conn = pyodbc.connect(conn_str)
        df = pl.read_database(query=f"SELECT * FROM {table}", connection=conn)
        return df
    except:
        print(TypeError)
    finally:
        conn.close()
    
def genera_concatenado(df):
    result = (df.with_columns(
                (pl.col("INTERMEDIARIO_ID") + pl.col("NUMERO_CREDITO"))
                .alias("Concatenado"))
            )
    return result

def genera_tpro_clave(df):
    result = df.with_columns(
        pl.when((pl.col("PROGRAMA_ID")>=32000)&(pl.col("PROGRAMA_ID")<=32100))
        .then(pl.col("PROGRAMA_ID"))
        .when((pl.col("PROGRAMA_ID")==3976)&(pl.col("PROGRAMA_ORIGINAL")==31415))
        .then(pl.col("PROGRAMA_ID"))
        .when((pl.col("PROGRAMA_ID")==33366)&(pl.col("PROGRAMA_ORIGINAL")==33842))
        .then(pl.col("PROGRAMA_ID"))
        .when((pl.col("PROGRAMA_ID").is_in([3536, 3537, 3539, 3542,3544, 3545, 3546,3547,3548,3549,3550, 3553, 3555, 3558,3559, 3560, 3564,3566]))&(pl.col("PROGRAMA_ORIGINAL")==3200))
        .then(pl.col("PROGRAMA_ID"))
        .when(pl.col("PROGRAMA_ORIGINAL")==3999)
        .then("PROGRAMA_ID")
        .otherwise(pl.col("PROGRAMA_ORIGINAL")).alias("TPRO_CLAVE")
    )
    return result