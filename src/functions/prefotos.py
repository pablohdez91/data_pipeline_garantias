# %%
import polars as pl
import pyodbc

# %%
def read_access(file, table):
    conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + file
    conn = pyodbc.connect(conn_str)

    df = pl.read_database(query=f"SELECT * FROM {table}", connection=conn)
    conn.close()
    return df

# %% [markdown]
# ### Querie Pagos

# %%
def import_f1():
    df_desembolsos = read_access(fl_desembolsos_p1, "DATOS")
    df_desembolsos_bmxt = read_access(fl_desembolsos_p1_bmxt, "DATOS")
    df_desembolsos_fianzas = read_access(fl_desembolsos_p1_fianzas, "DATOS")

    aux = pl.concat([
        df_desembolsos, 
        df_desembolsos_bmxt, 
        df_desembolsos_fianzas
        ], rechunk=True)
    
    df_pagos_f1 = (
        aux.select(
            pl.col("DESC_INDICADOR").alias("Producto"),
            pl.col("ESTATUS_RECUPERACION"),
            pl.col("FECHA_APERTURA").alias("Fecha de Apertura"),
            pl.col("FECHA_GARANTIA_HONRADA"),
            pl.col("FECHA_PRIMER_INCUMPLIMIENTO"),
            pl.col("FECHA_REGISTRO_ALTA").alias("Fecha Registro Alta"),
            pl.col("INTERMEDIARIO_ID"),
            pl.col("MONEDA_ID"),
            pl.col("NOMBRE_EMPRESA").alias("Empresa / Acreditado (Descripción)"),
            pl.col("NUMERO_CREDITO"),
            pl.col("PORCENTAJE_GARANTIZADO"),
            pl.col("PROGRAMA_ID"),
            pl.col("PROGRAMA_ORIGINAL"),
            pl.col("RAZON_SOCIAL").alias("Razón Social (Intermediario)"),
            pl.col("RFC_EMPRESA").alias("RFC Empresa / Acreditado"),
            pl.col("TIPO_CREDITO_ID"),
            pl.col("TIPO_GARANTIA_ID"),
            pl.col("TIPO_PERSONA"),
            pl.col("MONTO_CREDITO_MN (SUMA)").alias("Monto _Credito_Mn")
        )
    )
    return df_pagos_f1

def import_f2():
    df_desembolsos = read_access(fl_desembolsos_p2, "DATOS")
    df_desembolsos_bmxt = read_access(fl_desembolsos_p2_bmxt, "DATOS")
    df_desembolsos_fianzas = read_access(fl_desembolsos_p2_fianzas, "DATOS")

    aux = pl.concat([
        df_desembolsos, 
        df_desembolsos_bmxt, 
        df_desembolsos_fianzas
        ], rechunk=True)
    
    df_pagos_f2 = (
        aux.select(
            pl.col("DESC_INDICADOR").alias("Producto"),
            pl.col("FECHA_CONSULTA"),
            pl.col("FECHA_REGISTRO").alias("MIN Fecha_Registro"),
            pl.col("HISTORICO").alias("MAX Historico"),
            pl.col("INDICADOR_ID").alias("Producto ID"),
            pl.col("INTERMEDIARIO_ID"),
            pl.col("MONEDA_ID"),
            pl.col("NUMERO_CREDITO"),
            pl.col("PAGO_ID").alias("Pago ID"),
            pl.col("DFI_INTERESES_MORATORIOS (SUMA)").alias("SUM Intereses Moratorios"),
            pl.col("INTERES_DESEMBOLSO (SUMA)").alias("SUM Interes_Desembolso"),
            pl.col("MONTO_DESEMBOLSO (SUMA)").alias("SUM Monto_Desembolso")
        )
    )
    return df_pagos_f2



# %%
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


def genera_pagadas_global_inter(df1, df2):
    result = (df2.join(df1, on="Concatenado", how="left")
        .rename({
            "MIN Fecha_Registro": "MIN_Fecha_Registro",
            "SUM Monto_Desembolso":"Monto_Desembolsado",
            "SUM Interes_Desembolso": "Interes_Desembolso",
            "SUM Intereses Moratorios": "Interes_Moratorios"
        })
        .select([
            "Concatenado",
            "FECHA_CONSULTA", 
            "INTERMEDIARIO_ID", 
            "NUMERO_CREDITO", 
            "Producto", 
            "Pago ID",
            "Razón Social (Intermediario)",
            "MIN_Fecha_Registro",
            "FECHA_GARANTIA_HONRADA",
            "TPRO_CLAVE", 
            "PROGRAMA_ORIGINAL", 
            "PROGRAMA_ID", 
            "Monto _Credito_Mn", 
            "MONEDA_ID",
            "Fecha de Apertura",
            "TIPO_GARANTIA_ID",
            "TIPO_PERSONA",
            "RFC Empresa / Acreditado",
            "Monto_Desembolsado",
            "Interes_Desembolso",
            "Interes_Moratorios",
            "PORCENTAJE_GARANTIZADO",
            "TIPO_CREDITO_ID",
            "ESTATUS_RECUPERACION",
            "Empresa / Acreditado (Descripción)",
            "Fecha Registro Alta"
        ])
        .with_columns(pl.col("TIPO_GARANTIA_ID").fill_null(999))
    )
    return result

def genera_pagadas_global_vf(df):
    # Requiere que se hayan importado los catálogos
    result = (df
    .join(programa.select(['PROGRAMA_ID', 'AGRUPAMIENTO_ID', 'ESQUEMA', 'SUBESQUEMA']), on="PROGRAMA_ID", how='left')
        .join(agrupamiento, on='AGRUPAMIENTO_ID', how='left')
        .join(udis, left_on="Fecha de Apertura", right_on="Fecha_Paridad", how='left')
        .join(tipo_credito.select(['Tipo_Credito_ID', 'NR_R']), left_on='TIPO_CREDITO_ID', right_on="Tipo_Credito_ID", how='left')
        .join(tipo_garantia.select(['Tipo_garantia_ID', 'CSG']), left_on='TIPO_GARANTIA_ID', right_on='Tipo_garantia_ID', how='left')
        .join(sfc.select(['Intermediario_Id', 'CLAVE_CREDITO', 'FONDOS_CONTRAGARANTIA']), left_on=['INTERMEDIARIO_ID', 'NUMERO_CREDITO'], right_on=['Intermediario_Id', 'CLAVE_CREDITO'], how='left')
        )

    # Complementa
    result = (result
        .with_columns(pl.when(pl.col('MONEDA_ID')==54)
                    .then(tdc).otherwise(1).alias("TC"))
        .with_columns(pl.when(pl.col("Monto _Credito_Mn")<=(900000*pl.col("Paridad_Peso")))
                    .then(0).otherwise(1).alias("MM_UDIS"))
        .with_columns(pl.when(pl.col("FONDOS_CONTRAGARANTIA")=="SF")
                    .then(pl.lit("SF")).otherwise(pl.lit("CF")).alias("CSF"))
        )
    
    return result


def complementa_pagadas_global_vf(df):
    aux = df.select(['Monto_Desembolsado', 'Interes_Moratorios', 'Interes_Desembolso']).sum_horizontal(ignore_nulls=True)

    result = (df
        .with_columns((pl.col("Monto_Desembolsado") * pl.col("TC") * -1).alias("Monto_Desembolso_Mn"))
        .with_columns((pl.col("Interes_Desembolso") * pl.col("TC") * -1).alias("Interes_Desembolso_Mn"))
        .with_columns((pl.col("Interes_Moratorios") * pl.col("TC") * -1).alias("Interes_Moratorios_Mn"))
        .with_columns((aux).alias("Monto_Pagado_Mn"))
    )

    return result

def genera_validador(df):
    result = (df
        .group_by(["MAX Historico", "Producto", "MONEDA_ID"])
        .agg(pl.col("SUM Monto_Desembolso").sum(),
            pl.col("SUM Interes_Desembolso").sum(),
            pl.col("SUM Intereses Moratorios").sum()
            )
        )
    return result

def genera_pagadas_global_bancomext(df):
    result = (df.filter(
        (pl.col("Producto") == "GARANTIAS BANCOMEXT") |
        (pl.col("Producto") == "GARANTIAS SHF/LI FINANCIERO") |
        (pl.col("Producto") == "GARANTIAS BANSEFI") |
        (pl.col("Producto").is_null())
    ))
    return result

def genera_pagadas_global_sin_bancomext(df):
    result = (df.filter(
        (pl.col("Producto") != "GARANTIAS BANCOMEXT") &
        (pl.col("Producto") != "GARANTIAS SHF/LI FINANCIERO") &
        (pl.col("Producto") != "GARANTIAS BANSEFI")
    ))
    return result

def genera_valida_base_pagos_mn(pagadas_global_vf):
    valida_base_pagos_mn = (pagadas_global_vf
    .group_by('Producto')
    .agg(pl.col('Monto_Desembolso_Mn').sum().alias('Monto_Desembolso_Mn_Suma'),
        pl.col('Interes_Desembolso_Mn').sum().alias('Interes_Desembolso_Mn_Suma'),
        pl.col('Interes_Moratorios_Mn').sum().alias('Interes_Moratorios_Mn_Sum')
        )
    )
    return valida_base_pagos_mn

def genera_valida_base_pagos(pagadas_global_vf):
     valida_base_pagos = (pagadas_global_vf
     .group_by('Producto')
     .agg(pl.col('Monto_Desembolsado').sum().alias('Monto_Desembolsado_Suma'),
          pl.col('Interes_Desembolso').sum().alias('Interes_Desembolso_Suma'),
          pl.col('Interes_Moratorios').sum().alias('Interes_Moratorios_Sum')
          )
     )
     return valida_base_pagos

# %% [markdown]
# #### Querie Recuperadas

# %%
def importa_recuperaciones():
    schema_recuperaciones = {
        'ANIO DEGL': pl.String,
        'DESC_INDICADOR': pl.String,
        'DESCRIPCION': pl.String,
        'ESTATUS': pl.String,
        'FECHA': pl.Datetime,
        'FECHA_APERTURA': pl.Datetime,
        'FECHA_CONSULTA': pl.Datetime,
        'FECHA_GARANTIA_HONRADA': pl.Datetime,
        'FECHA_REGISTRO': pl.Datetime,
        'FECHA_REGISTRO_ALTA': pl.Datetime,
        'FISO_ID': pl.Int32,
        'HISTORICO': pl.String,
        'ID': pl.Int64,
        'INTERMEDIARIO_ID': pl.String,
        'MES DEGL': pl.Int8,
        'MONEDA_ID': pl.Int64,
        'NOMBRE_EMPRESA': pl.String,
        'NUMERO_CREDITO': pl.String,
        'PORCENTAJE_GARANTIZADO': pl.Float32,
        'PROGRAMA_ID': pl.Int32,
        'PROGRAMA_ORIGINAL': pl.Int32,
        'RAZON_SOCIAL': pl.String,
        'RFC_EMPRESA': pl.String,
        'TIPO_CAMBIO_GARANTIA': pl.Float32,
        'TIPO_CREDITO_ID': pl.Int32,
        'TIPO_GARANTIA_ID': pl.Int32,
        'TIPO_PERSONA': pl.String,
        'GASTO_JUICIOS (SUMA)': pl.Float32,
        'INTER_MORAT (SUMA)':pl.Float32,
        'INTERES_GENERADO (SUMA)':pl.Float32,
        'INTERESES (SUMA)': pl.Float32,
        'MONTO (SUMA)': pl.Float32,
        'MONTO_CREDITO_MN (SUMA)': pl.Float64,
        'MORATORIOS (SUMA)': pl.Float32,
        'Número de registros': pl.Int8,
        'PENALIZACION (SUMA)': pl.Float32
    }

    df_dwh_recup = (pl.read_csv(fl_recuperaciones, schema=schema_recuperaciones)
                    .drop('Número de registros'))
    df_dwh_recup = df_dwh_recup.drop(['ANIO DEGL', 'FISO_ID', 'MES DEGL'])

    [schema_recuperaciones.pop(key) for key in ["Número de registros", "ANIO DEGL", "FISO_ID", "MES DEGL"]]

    df_dwh_recup_bmxt = read_access(fl_recuperaciones_bmxt, "DATOS")
    

    aux = pl.concat([
        df_dwh_recup, 
        df_dwh_recup_bmxt.cast(schema_recuperaciones)
        ], rechunk=True)

    df_dwh_recuperaciones = aux.rename({
        "DESC_INDICADOR": "Producto",
        "FECHA_REGISTRO_ALTA": "Fecha Registro Alta",
        "NOMBRE_EMPRESA": "Empresa / Acreditado (Descripción)",
        "RAZON_SOCIAL": "Razón Social (Intermediario)",
        "RFC_EMPRESA": "RFC Empresa / Acreditado",
        "TIPO_CAMBIO_GARANTIA":"Tipo_Cambio_Cierre",
        "GASTO_JUICIOS (SUMA)": "Gastos Juicio",
        "INTER_MORAT (SUMA)": "Moratorios",
        "INTERES_GENERADO (SUMA)": "Interes Generado", 
        "INTERESES (SUMA)": "Interes", 
        "MONTO (SUMA)": "Monto",
        "MONTO_CREDITO_MN (SUMA)": "Monto _Credito_Mn",
        "MORATORIOS (SUMA)": "Excedente",
        "PENALIZACION (SUMA)": "Penalizacion"
    }).select([
        "Producto", 
        "DESCRIPCION", 
        "ESTATUS", 
        "FECHA", 
        "FECHA_APERTURA", 
        "FECHA_CONSULTA", 
        "FECHA_GARANTIA_HONRADA", 
        "FECHA_REGISTRO",
        "Fecha Registro Alta",
        "HISTORICO", 
        "ID", 
        "INTERMEDIARIO_ID", 
        "MONEDA_ID", 
        "Empresa / Acreditado (Descripción)", 
        "NUMERO_CREDITO", 
        "PORCENTAJE_GARANTIZADO", 
        "PROGRAMA_ID", 
        "PROGRAMA_ORIGINAL", 
        "Razón Social (Intermediario)", 
        "RFC Empresa / Acreditado",
        "Tipo_Cambio_Cierre", 
        "TIPO_CREDITO_ID", 
        "TIPO_GARANTIA_ID", 
        "TIPO_PERSONA", 
        "Gastos Juicio", 
        "Moratorios", 
        "Interes Generado", 
        "Interes", 
        "Monto", 
        "Monto _Credito_Mn", 
        "Excedente",
        "Penalizacion"
    ])

    df_dwh_recuperaciones = df_dwh_recuperaciones.with_columns(pl.col("TIPO_GARANTIA_ID").fill_null(999))
    df_dwh_recuperaciones = genera_tpro_clave(df_dwh_recuperaciones)

    return df_dwh_recuperaciones

# %%
def genera_recuperadas_global_inter(df):
    # Requiere que se hayan importado los catálogos
    result = (df
    .join(programa.select(['PROGRAMA_ID', 'AGRUPAMIENTO_ID', 'ESQUEMA', 'SUBESQUEMA']), on="PROGRAMA_ID", how='left')
        .join(agrupamiento, on='AGRUPAMIENTO_ID', how='left')
        .join(udis, left_on="FECHA_APERTURA", right_on="Fecha_Paridad", how='left')
        .join(tipo_credito.select(['Tipo_Credito_ID', 'NR_R']), left_on='TIPO_CREDITO_ID', right_on="Tipo_Credito_ID", how='left')
        .join(tipo_garantia.select(['Tipo_garantia_ID', 'CSG']), left_on='TIPO_GARANTIA_ID', right_on='Tipo_garantia_ID', how='left')
        .join(sfc.select(['Intermediario_Id', 'CLAVE_CREDITO', 'FONDOS_CONTRAGARANTIA']), left_on=['INTERMEDIARIO_ID', 'NUMERO_CREDITO'], right_on=['Intermediario_Id', 'CLAVE_CREDITO'], how='left')
        .join(estatus.select(['Estatus ID', 'Recup/Rescat']), left_on='ESTATUS', right_on='Estatus ID', how='left')
        ).rename({'Recup/Rescat': 'Recup_Rescat'})

    # Complementa
    result = (result
        .with_columns(pl.when(pl.col('MONEDA_ID')==54)
                    .then(tdc).otherwise(1).alias("TC"))
        .with_columns(pl.when(pl.col("Monto _Credito_Mn")<=(900000*pl.col("Paridad_Peso")))
                    .then(0).otherwise(1).alias("MM_UDIS"))
        .with_columns(pl.when(pl.col("FONDOS_CONTRAGARANTIA")=="SF")
                    .then(pl.lit("SF")).otherwise(pl.lit("CF")).alias("CSF"))
        )
    
    return result

def genera_recuperadas_global_vf(df):
    aux = df.select(['Monto', 'Interes', 'Moratorios', 'Excedente']).sum_horizontal(ignore_nulls=True)
    result = (df
        .with_columns((pl.col("Monto") * pl.col("TC")).alias("Monto_Mn"))
        .with_columns((pl.col("Interes") * pl.col("TC")).alias("Interes_Mn"))
        .with_columns((pl.col("Moratorios") * pl.col("TC")).alias("Moratorios_Mn"))
        .with_columns((pl.col("Excedente") * pl.col("TC")).alias("Excedente_Mn"))
        .with_columns((pl.col("Gastos Juicio") * pl.col("TC")).alias("Gastos_Juicio_Mn"))
        .with_columns((aux * pl.col("TC")).alias("Sub_Total_Mn"))
        .with_columns((pl.col("Sub_Total_Mn")-pl.col("Gastos_Juicio_Mn")).alias("Monto_Total_Mn"))
        .with_columns(pl.when(((pl.col("NUMERO_CREDITO")=="9842725312") & (pl.col("INTERMEDIARIO_ID")=="10040012")))
               .then(date(2011,8,2))
               .otherwise(pl.col("FECHA_GARANTIA_HONRADA"))
               .alias("FECHA_GARANTIA_HONRADA"))
    )
    return result

def genera_recuperadas_global_bancomext(df):
    result = (df.filter(
        (pl.col("Producto") == "GARANTIAS BANCOMEXT") |
        (pl.col("Producto") == "GARANTIAS SHF/LI FINANCIERO") |
        (pl.col("Producto") == "GARANTIAS BANSEFI") |
        (pl.col("Producto").is_null())
    ))
    return result

def genera_recuperadas_global_sin_bancomext(df):
    result = (df.filter(
        (pl.col("Producto") != "GARANTIAS BANCOMEXT") &
        (pl.col("Producto") != "GARANTIAS SHF/LI FINANCIERO") &
        (pl.col("Producto") != "GARANTIAS BANSEFI")
    ))
    return result

def genera_valida_dwh_dac(df):
    result = (df
      .group_by(['HISTORICO', 'Tipo_Cambio_Cierre', 'Producto'])
      .agg(pl.col("Monto").sum().alias("S_Monto"),
            pl.col("Interes").sum().alias("S_Interes"),
            pl.col("Moratorios").sum().alias("S_Moratorios"),
            pl.col("Excedente").sum().alias("S_Excedente"),
            pl.col("Gastos Juicio").sum().alias("S_Gastos_Juicio"),
          )
    )
    return result

def genera_valida_td(df):
    result = (df
      .group_by(['Producto', 'Tipo_Cambio_Cierre'])
      .agg(pl.col("Monto").sum().alias("S_Monto"),
            pl.col("Interes").sum().alias("S_Interes"),
            pl.col("Moratorios").sum().alias("S_Moratorios"),
            pl.col("Excedente").sum().alias("S_Excedente"),
            pl.col("Gastos Juicio").sum().alias("S_Gastos_Juicio"),
          )
      )
    return result

# %%


# %% [markdown]
# #### Querie UnionFlujos

# %%
def genera_pagos_agrup(pagadas_global_vf):
    pagos_agrup = pagadas_global_vf.group_by([
        'Producto', 
        'INTERMEDIARIO_ID', 
        'NUMERO_CREDITO', 
        'FECHA_GARANTIA_HONRADA', 
        'MM_UDIS', 
        'NR_R', 
        'Razón Social (Intermediario)', 
        'TPRO_CLAVE', 
        'AGRUPAMIENTO'
    ]).agg(
        pl.col("Monto_Desembolso_Mn").sum().alias("Monto_Desem_Mn"),
        pl.col("Interes_Desembolso_Mn").sum().alias("Interes_Desem_Mn"),
        pl.col("Interes_Moratorios_Mn").sum().alias("Interes_Morat_Mn"),
        pl.col("Monto_Pagado_Mn").sum().alias("MPagado_Mn"),
    )


    return pagos_agrup

def genera_recuperaciones_agrup(recuperadas_global_vf):
    recuperaciones_agrup = (recuperadas_global_vf
        .group_by(["ESTATUS", "INTERMEDIARIO_ID", "NUMERO_CREDITO", "Concatenado"])
        .agg(pl.col("Monto_Mn").sum().alias("Monto_Recup_Mn"),
            pl.col("Interes_Mn").sum().alias("Interes_Recup_Mn"),
            pl.col("Moratorios_Mn").sum().alias("Moratorios_Recup_Mn"),
            pl.col("Excedente_Mn").sum().alias("Excedente_Recup_Mn"),
            pl.col("Monto_Total_Mn").sum().alias("Monto_Total_Recup_Mn"))
        )
    return recuperaciones_agrup

def genera_recupera_con_pagos_flujos(uf_recuperaciones_pagos):
    order_columns = [
        "Concatenado", 
        "FECHA_CONSULTA", 
        "PROGRAMA_ID",
        "TIPO_GARANTIA_ID", 
        "Tipo_Cambio_Cierre", 
        "PROGRAMA_ORIGINAL", 
        "PORCENTAJE_GARANTIZADO",
        "Monto _Credito_Mn", 
        "FECHA_APERTURA", 
        "MONEDA_ID", 
        "TIPO_CREDITO_ID", 
        "FECHA_GARANTIA_HONRADA", 
        "TPRO_CLAVE",
        "NR_R", 
        "Producto", 
        "AGRUPAMIENTO_ID", 
        "AGRUPAMIENTO", 
        "INTERMEDIARIO_ID",
        "NUMERO_CREDITO", 
        "ID", 
        "Monto", 
        "Interes", 
        "Moratorios", 
        "DESCRIPCION", 
        "ESTATUS", 
        "FECHA_REGISTRO", 
        "FECHA", 
        "Monto_Mn",
        "Interes_Mn", 
        "Excedente", 
        "Excedente_Mn", 
        "Moratorios_Mn", 
        "Sub_Total_Mn",
        "Razón Social (Intermediario)",
        "Empresa / Acreditado (Descripción)",
        "RFC Empresa / Acreditado", 
        "TIPO_PERSONA", 
        "Recup_Rescat",
        "MM_UDIS", 
        "ESQUEMA", 
        "CSG", 
        "CSF", 
        "Gastos_Juicio_Mn", 
        "Monto_Total_Mn", 
        "MPagado_Mn",
        "HISTORICO", 
        "Fecha Registro Alta"
        ]

    recupera_con_pagos_flujos = uf_recuperaciones_pagos.select(order_columns)
    return recupera_con_pagos_flujos

# %%
def genera_recupera_con_pagos_flujos_bancomext(df):
    result = (df.filter(
        (pl.col("Producto") == "GARANTIAS BANCOMEXT") |
        (pl.col("Producto") == "GARANTIAS SHF/LI FINANCIERO") |
        (pl.col("Producto") == "GARANTIAS BANSEFI") |
        (pl.col("Producto").is_null())
    ))
    return result


def genera_recupera_con_pagos_flujos_sin_bancomext(df):
    result = (df.filter(
        (pl.col("Producto") != "GARANTIAS BANCOMEXT") |
        (pl.col("Producto") != "GARANTIAS SHF/LI FINANCIERO") |
        (pl.col("Producto") != "GARANTIAS BANSEFI")
    ))
    return result

def genera_valida_base_pagos_recup_mn(recupera_con_pagos_flujos):
     valida_base_pagos_mn = (recupera_con_pagos_flujos
     .group_by('Producto')
     .agg(pl.col('Monto_Mn').sum().alias('Monto_Mn_Suma'),
          pl.col('Interes_Mn').sum().alias('Interes_Mn_Suma'),
          pl.col('Moratorios_Mn').sum().alias('Moratorios_Mn_Suma'),
          pl.col('Excedente_Mn').sum().alias('Excedente_Mn_Suma'),
          pl.col('Gastos_Juicio_Mn').sum().alias('Gastos_Juicio_Mn_Suma')
          )
     )
     return valida_base_pagos_mn

def genera_valida_base_pagos_recup(recupera_con_pagos_flujos):
     valida_base_pagos = (recupera_con_pagos_flujos
     .group_by('Producto')
     .agg(pl.col('Monto').sum().alias('Monto_Suma'),
          pl.col('Interes').sum().alias('Interes_Suma'),
          pl.col('Moratorios').sum().alias('Moratorios_Suma'),
          pl.col('Excedente').sum().alias('Excedente_Suma'),
          pl.col('Gastos Juicio').sum().alias('Gastos_Juicio_Suma')
          )
     )
     return valida_base_pagos

# %%



