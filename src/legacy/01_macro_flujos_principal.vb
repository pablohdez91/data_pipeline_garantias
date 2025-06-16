Private Sub Principal()

Dim dbs As DAO.Database
mes_num = 202409
anio = Left(mes_num, 4)
mes_1 = IIf(Mid(mes_num, 5, 1) = 0, Right(mes_num, 1), Right(mes_num, 2))
mes = mes_letra(mes_1) & Mid(mes_num, 3, 2)

' Prefixes
'   wd: Working Directory
'   db: Database (Access file)
'   tbl: Table

wd = "D:\DAR\proyecto_mejora_fotos\2. Nuevas fotos\"
wd_external = wd & "data\external\"
wd_processed = wd & "data\processed\"
wd_processed_dwh = wd_processed & "DWH\"
wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
wd_raw = wd & "data\raw\"
wd_staging = wd & "data\staging\"
wd_validations = wd & "data\validations\"

'BDs Inputs
db_desembolsos_p1 = wd_raw & "Desembolsos_P1_" & mes_num & ".mdb"
db_desembolsos_p1_bmxt = wd_raw & "Desembolsos_P1_BMXT_" & mes_num & ".mdb"
db_desembolsos_p1_80686 = wd_raw & "Desembolsos_P1_80686_" & mes_num & ".mdb"
db_desembolsos_p2 = wd_raw & "Desembolsos_P2_" & mes_num & ".mdb"
db_desembolsos_p2_bmxt = wd_raw & "Desembolsos_P2_BMXT_" & mes_num & ".mdb"
db_desembolsos_p2_80686 = wd_raw & "Desembolsos_P2_80686_" & mes_num & ".mdb"
db_recuperaciones = wd_raw & "Recuperaciones_" & mes_num & ".mdb"
db_recuperaciones_bmxt = wd_raw & "Recuperaciones_BMXT_" & mes_num & ".mdb"

'BDs Outputs
db_catalogos = wd_external & "Catálogos_" & mes & ".accdb"
db_base = wd_staging & "Base.accdb"
db_base_vacia = wd_staging & "Base_Vacia.accdb"
db_querie_pagadas = wd_processed_dwh & "Querie_Pagadas_" & mes_num & ".accdb"
db_querie_recuperaciones = wd_processed_dwh & "Querie_Recuperaciones_" & mes_num & ".accdb"
db_querie_union_flujos = wd_processed_dwh & "Querie_UnionFlujos_" & mes_num & ".accdb"

db_pagadas_global_finales = wd_processed_dwh_bases_finales & "Pagadas_Global_VF_" & mes_num & ".accdb"
db_recupera_con_pagos_flujos_finales = wd_processed_dwh_bases_finales & "Recupera_con_Pagos_Flujos_" & mes_num & ".accdb"
db_pagadas_global_finales_xl = wd_processed_dwh_bases_finales & "Pagadas_Global_VF_" & mes_num & ".xlsx"
db_recupera_con_pagos_flujos_finales_xl = wd_processed_dwh_bases_finales & "Recupera_con_Pagos_Flujos_" & mes_num & ".xlsx"

db_pagadas_global_entregables = wd_processed_dwh_entregables & "Pagadas_Global_VF_" & mes_num & ".accdb"
db_recupera_con_pagos_flujos_entregables = wd_processed_dwh_entregables & "Recupera_con_Pagos_Flujos_" & mes_num & ".accdb"
db_pagadas_global_entregables_xl = wd_processed_dwh_entregables & "Pagadas_Global_VF_" & mes_num & ".xlsx"
db_recupera_con_pagos_flujos_entregables_xl = wd_processed_dwh_entregables & "Recupera_con_Pagos_Flujos_" & mes_num & ".xlsx"

'Tables
tbl_dwh_pagos_f1 = "DWH_Pagos_F1_" & mes_num
tbl_dwh_pagos_f2 = "DWH_Pagos_F2_" & mes_num
tbl_pagadas_global_vf = "Pagadas_Global_VF_" & mes_num
tbl_pagadas_global_vf_inter = tbl_pagadas_global_vf & "_Inter"
tbl_dwh_recuperaciones = "DWH_Recuperaciones_" & mes_num
tbl_recuperadas_global_vf = "Recuperadas_Global_VF_" & mes_num
tbl_recuperadas_valida_dwh_dac = "Recuperadas_Valida_DWHvsDAC"
tbl_recuperadas_valida_td = "Recuperadas_Valida_TD"
tbl_recuperadas_global_vf_inter = tbl_recuperadas_global_vf & "_Inter"
tbl_pagos_agrup = "T_Pagos_Agrup_" & mes_num
tbl_recuperaciones_agrup = "T_UF_Recuperaciones_Agrup_" & mes_num
tbl_uf_pagos_recuperaciones = "T_UF_Pagos-Recuperaciones_" & mes_num
tbl_uf_recuperaciones_pagos = "T_UF_Recuperaciones-Pagos" & mes_num
tbl_roberto = "T_Roberto_" & mes_num
tbl_recupera_con_pagos_flujos = "Recupera_con_Pagos_Flujos_" & mes_num
tbl_recupera_con_pagos_flujos_ord = "Recupera_con_Pagos_Flujos_" & mes_num & "_Ord"
tbl_pagos_p1 = "TBL_PagosP1_" & mes_num
tbl_pagos_p1_bmxt = "TBL_PagosP1_BMXT_" & mes_num
tbl_pagos_p1_80686 = "TBL_PagosP1_80686_" & mes_num
tbl_pagos_p2 = "TBL_PagosP2_" & mes_num
tbl_pagos_p2_bmxt = "TBL_PagosP2_BMXT_" & mes_num
tbl_pagos_p2_80686 = "TBL_PagosP2_80686_" & mes_num
tbl_recup = "TBL_Recup_" & mes_num
tbl_recup_bmxt = "TBL_Recup_BMXT_" & mes_num


If existe_ruta(wd_processed_dwh) = False Then
    MkDir (wd_processed_dwh)
End If

If existe_ruta(db_querie_pagadas) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh & "Copia de Seguridad.accdb")
    Name wd_processed_dwh & "Copia de Seguridad.accdb" As db_querie_pagadas
End If

If existe_ruta(db_querie_recuperaciones) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh & "Copia de Seguridad.accdb")
    Name wd_processed_dwh & "Copia de Seguridad.accdb" As db_querie_recuperaciones
End If
'Crea Base_UnionFlujos
If existe_ruta(db_querie_union_flujos) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh & "Copia de Seguridad.accdb")
    Name wd_processed_dwh & "Copia de Seguridad.accdb" As db_querie_union_flujos
End If


'Vicula Catalogos a Pagos
Call vincula_tabla(db_catalogos, db_querie_pagadas, "TIPO CAMBIO", "TIPO CAMBIO") 'Vincula TC
Call vincula_tabla(db_catalogos, db_querie_pagadas, "PROGRAMA", "PROGRAMA") 'Vincula Programa
Call vincula_tabla(db_catalogos, db_querie_pagadas, "AGRUPAMIENTO", "AGRUPAMIENTO") 'Vincula Primeras_Perdidas
Call vincula_tabla(db_catalogos, db_querie_pagadas, "UDIS", "UDIS") 'Vincula UDIS
Call vincula_tabla(db_catalogos, db_querie_pagadas, "TIPO_CREDITO", "TIPO_CREDITO") 'Vincula Tipo_Crédito_Id
Call vincula_tabla(db_catalogos, db_querie_pagadas, "TIPO_GARANTIA", "TIPO_GARANTIA") 'Vincula Tipo_Garantía_Id
Call vincula_tabla(db_catalogos, db_querie_pagadas, "SIN FONDOS CONTRAGARANTIA", "SIN FONDOS CONTRAGARANTIA") 'Vincula SIN FONDOS CONTRAGARANTIA

'ACOV 201805 cambio Tableau
'Vinculo las bases de Pagos
Call vincula_tabla(db_desembolsos_p1, db_querie_pagadas, "DATOS", tbl_pagos_p1)
Call vincula_tabla(db_desembolsos_p2, db_querie_pagadas, "DATOS", tbl_pagos_p2)
Call vincula_tabla(db_desembolsos_p1_bmxt, db_querie_pagadas, "DATOS", tbl_pagos_p1_bmxt)
Call vincula_tabla(db_desembolsos_p2_bmxt, db_querie_pagadas, "DATOS", tbl_pagos_p2_bmxt)
Call vincula_tabla(db_desembolsos_p1_80686, db_querie_pagadas, "DATOS", tbl_pagos_p1_80686)
Call vincula_tabla(db_desembolsos_p2_80686, db_querie_pagadas, "DATOS", tbl_pagos_p2_80686)

'Vicula Catalogos a Recuperaciones
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "ESTATUS", "ESTATUS") 'Vincula Estatus
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "Tipo Cambio", "Tipo Cambio") 'Vincula TC
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "PROGRAMA", "PROGRAMA") 'Vincula Primeras_Perdidas
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "AGRUPAMIENTO", "AGRUPAMIENTO") 'Vincula Primeras_Perdidas
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "UDIS", "UDIS") 'Vincula UDIS
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "TIPO_CREDITO", "TIPO_CREDITO") 'Vincula Tipo_Crédito_Id
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "TIPO_GARANTIA", "TIPO_GARANTIA") 'Vincula Tipo_Garantía_Id
Call vincula_tabla(db_catalogos, db_querie_recuperaciones, "SIN FONDOS CONTRAGARANTIA", "SIN FONDOS CONTRAGARANTIA") 'Vincula SIN FONDOS CONTRAGARANTIA

'ACOV 201805 cambio Tableau
'Vinculo las bases de recuperaciones
Call vincula_tabla(db_recuperaciones, db_querie_recuperaciones, "DATOS", tbl_recup)
Call vincula_tabla(db_recuperaciones_bmxt, db_querie_recuperaciones, "DATOS", tbl_recup_bmxt)


'Empiezo proceso
Set dbs = OpenDatabase(db_querie_pagadas)
    'ACOV 201805 TBL Uno la base de BMXT y la sBMXT
    dbs.Execute "SELECT * INTO AUX_P1 FROM " & tbl_pagos_p1 & ";"
    dbs.Execute "SELECT * INTO AUX_P2 FROM " & tbl_pagos_p2 & ";"
    Call inserta_filas_in(dbs, "", tbl_pagos_p1_bmxt, "AUX_P1")
    Call inserta_filas_in(dbs, "", tbl_pagos_p2_bmxt, "AUX_P2")
    Call inserta_filas_in(dbs, "", tbl_pagos_p1_80686, "AUX_P1")
    Call inserta_filas_in(dbs, "", tbl_pagos_p2_80686, "AUX_P2")
    
    dbs.Execute "SELECT DESC_INDICADOR AS Producto, ESTATUS_RECUPERACION, FECHA_APERTURA AS [Fecha de Apertura], FECHA_GARANTIA_HONRADA, " _
        & " FECHA_PRIMER_INCUMPLIMIENTO, FECHA_REGISTRO_ALTA AS [Fecha Registro Alta], INTERMEDIARIO_ID, MONEDA_ID, NOMBRE_EMPRESA AS [Empresa / Acreditado (Descripción)], " _
        & " NUMERO_CREDITO, PORCENTAJE_GARANTIZADO, PROGRAMA_ID, PROGRAMA_ORIGINAL, RAZON_SOCIAL AS [Razón Social (Intermediario)], " _
        & " RFC_EMPRESA AS [RFC Empresa / Acreditado], TIPO_CREDITO_ID, TIPO_GARANTIA_ID, TIPO_PERSONA, [MONTO_CREDITO_MN (SUMA)] AS [Monto _Credito_Mn] " _
        & " INTO " & tbl_dwh_pagos_f1 & " FROM AUX_P1 ;"
     
    dbs.Execute "SELECT DESC_INDICADOR AS Producto, FECHA_CONSULTA, FECHA_REGISTRO AS [MIN Fecha_Registro], HISTORICO AS [MAX Historico], " _
        & " INDICADOR_ID AS [Producto ID], INTERMEDIARIO_ID, MONEDA_ID, NUMERO_CREDITO, PAGO_ID AS [Pago ID], " _
        & " [DFI_INTERESES_MORATORIOS (SUMA)] AS [SUM Intereses Moratorios], [INTERES_DESEMBOLSO (SUMA)] AS [SUM Interes_Desembolso], " _
        & " [MONTO_DESEMBOLSO (SUMA)] AS [SUM Monto_Desembolso] INTO " & tbl_dwh_pagos_f2 & " FROM AUX_P2;"
    
    'Inserta en la base P1 el crédito faltante en DWH
   ' Call inserta_filas_in(dbs, "", "BD_DWH_Pagos_F1_EmpComp", tbl_dwh_pagos_f1)
    
    'Borro las bases auxiliares
    dbs.Execute "DROP TABLE AUX_P1;"
    dbs.Execute "DROP TABLE AUX_P2;"
    
    'Proceso anterior
    Call corrige_campos(dbs, tbl_dwh_pagos_f1, "Producto", " IIF((Producto = 'GARANTIA SUBASTA' and  (Programa_Original=31768 and Programa_Id=34008)), 'GARANTIA EMPRESARIAL', Producto) ", "")
    Call corrige_campos(dbs, tbl_dwh_pagos_f1, "Producto", " IIF((Producto = 'GARANTIA SUBASTA' and  (Programa_Original=31437 and Programa_Id=34007)), 'GARANTIA EMPRESARIAL', Producto) ", "")
    Call inserta_columna(dbs, tbl_dwh_pagos_f1, "Concatenado_P1", "Text(200)")
    Call corrige_campos(dbs, tbl_dwh_pagos_f1, "Concatenado_P1", "A.[Intermediario_Id]&A.[Numero_Credito]", "")
    Call inserta_columna(dbs, tbl_dwh_pagos_f2, "Concatenado_P2", "Text(200)")
    Call corrige_campos(dbs, tbl_dwh_pagos_f2, "Concatenado_P2", "A.[Intermediario_Id]&A.[Numero_Credito]", "")
    
    'ACOV 201805 agregamos la base de empresa mediana
    Call inserta_columna(dbs, tbl_dwh_pagos_f1, "TPRO_CLAVE", "Double")
    Call corrige_campos(dbs, tbl_dwh_pagos_f1, "TPRO_CLAVE", "IIf(A.Programa_Id>=32000 And A.Programa_Id<=32100, A.Programa_Id, IIf(A.Programa_Id=3976 And A.Programa_Original=31415,A.Programa_Id,IIF(A.Programa_Original = 33842 AND A.Programa_Id = 33366, A.Programa_Id, IIF(A.Programa_Original = 3200 AND A.Programa_Id IN (3536, 3537, 3539, 3542,3544, 3545, 3546,3547,3548,3549,3550, 3553, 3555, 3558,3559, 3560, 3564,3566), A.Programa_Id,IIf(A.Programa_Original = 3999,A.Programa_Id,A.Programa_Original))))) ", "")
    
    Call cruza_pagof1_pagof2(dbs, tbl_pagadas_global_vf_inter, tbl_dwh_pagos_f1, tbl_dwh_pagos_f2)
    Call corrige_campos(dbs, tbl_pagadas_global_vf_inter, "Tipo_Garantia_Id", 999, "Where A.Tipo_Garantia_Id is null ")
    Call cruza_catalogos_2(dbs, tbl_pagadas_global_vf, tbl_pagadas_global_vf_inter, "Tipo Cambio", "PROGRAMA", "UDIS", "AGRUPAMIENTO", "TIPO_CREDITO", "TIPO_GARANTIA", "Fecha de Apertura", "SIN FONDOS CONTRAGARANTIA")
     
     
     'Agrega Montos en MN
    Call inserta_columna(dbs, tbl_pagadas_global_vf, "Monto_Desembolso_Mn", "Double")
    Call corrige_campos(dbs, tbl_pagadas_global_vf, "Monto_Desembolso_Mn", "A.[Monto_Desembolsado]*-1*A.TC", "")
    Call inserta_columna(dbs, tbl_pagadas_global_vf, "Interes_Desembolso_Mn", "Double")
    Call corrige_campos(dbs, tbl_pagadas_global_vf, "Interes_Desembolso_Mn", "A.[Interes_Desembolso]*-1*A.TC", "")
    Call inserta_columna(dbs, tbl_pagadas_global_vf, "Interes_Moratorios_Mn", "Double")
    Call corrige_campos(dbs, tbl_pagadas_global_vf, "Interes_Moratorios_Mn", "A.[Interes_Moratorios]*-1*A.TC", "")
    Call inserta_columna(dbs, tbl_pagadas_global_vf, "Monto_Pagado_Mn", "Double")
    Call corrige_campos(dbs, tbl_pagadas_global_vf, "Monto_Pagado_Mn", "(IIF(A.[Monto_Desembolsado] is null,0,A.[Monto_Desembolsado])+IIF(A.[Interes_Desembolso] is null,0,A.[Interes_Desembolso])+IIF(A.[Interes_Moratorios] is null,0,A.[Interes_Moratorios]))*A.TC*-1", "")
    Call extracto_bd(dbs, tbl_pagadas_global_vf, tbl_pagadas_global_vf & "_sBancomext", " WHERE Producto <> 'GARANTIAS BANCOMEXT' or Producto <>'GARANTIAS SHF/LI FINANCIERO' or Producto <> 'GARANTIAS BANSEFI' ")
    Call extracto_bd(dbs, tbl_pagadas_global_vf, tbl_pagadas_global_vf & "_Bancomext", " WHERE (Producto = 'GARANTIAS BANCOMEXT' or Producto ='GARANTIAS SHF/LI FINANCIERO' or Producto is null or Producto = 'GARANTIAS BANSEFI') ")
dbs.Close

'Importa Recuperaciones 467857270
Set dbs = OpenDatabase(db_querie_recuperaciones)
    
    'ACOV 201805 cambio a TBL
    dbs.Execute "SELECT * INTO AUX FROM " & tbl_recup & ";"
    Call inserta_filas_in(dbs, "", tbl_recup_bmxt, "AUX")
    
    dbs.Execute "SELECT DESC_INDICADOR AS Producto, DESCRIPCION, ESTATUS, FECHA, FECHA_APERTURA, FECHA_CONSULTA, FECHA_GARANTIA_HONRADA, FECHA_REGISTRO, " _
        & " FECHA_REGISTRO_ALTA as [Fecha Registro Alta], HISTORICO, ID, INTERMEDIARIO_ID, MONEDA_ID, NOMBRE_EMPRESA as [Empresa / Acreditado (Descripción)], " _
        & " NUMERO_CREDITO, PORCENTAJE_GARANTIZADO, PROGRAMA_ID, PROGRAMA_ORIGINAL, RAZON_SOCIAL as [Razón Social (Intermediario)], RFC_EMPRESA AS [RFC Empresa / Acreditado], " _
        & " TIPO_CAMBIO_GARANTIA AS [Tipo_Cambio_Cierre], TIPO_CREDITO_ID, TIPO_GARANTIA_ID, TIPO_PERSONA, [GASTO_JUICIOS (SUMA)] AS [Gastos Juicio], " _
        & " [INTER_MORAT (SUMA)] AS Moratorios, [INTERES_GENERADO (SUMA)] as [Interes Generado],  " _
        & " [INTERESES (SUMA)] AS Interes, [MONTO (SUMA)] AS Monto, [MONTO_CREDITO_MN (SUMA)] AS [Monto _Credito_Mn], [MORATORIOS (SUMA)] AS Excedente, " _
        & " [PENALIZACION (SUMA)] AS Penalizacion " _
        & " INTO " & tbl_dwh_recuperaciones & " FROM AUX;"
        
    Call inserta_columna(dbs, tbl_dwh_recuperaciones, "TPRO_CLAVE", "Double")
    'Call corrige_campos(dbs, tbl_dwh_recuperaciones, "TPRO_CLAVE", "IIf(A.Programa_Id>=32000 And A.Programa_Id<=32100, A.Programa_Id, IIf(A.Programa_Id=3976 And A.Programa_Original=31415,A.Programa_Id,IIf(A.Programa_Original = 3999,A.Programa_Id,A.Programa_Original))) ", "")
    Call corrige_campos(dbs, tbl_dwh_recuperaciones, "TPRO_CLAVE", "IIf(A.Programa_Id>=32000 And A.Programa_Id<=32100, A.Programa_Id, IIf(A.Programa_Id=3976 And A.Programa_Original=31415,A.Programa_Id,IIF(A.Programa_Original = 33842 AND A.Programa_Id = 33366, A.Programa_Id, IIF(A.Programa_Original = 3200 AND A.Programa_Id IN (3536, 3537, 3539, 3542,3544, 3545, 3546,3547,3548,3549,3550, 3553, 3555, 3558,3559, 3560, 3564,3566), A.Programa_Id,IIf(A.Programa_Original = 3999,A.Programa_Id,A.Programa_Original))))) ", "")
    dbs.Execute "DROP TABLE AUX"
    
    'Proceso viejo
    Call corrige_campos(dbs, tbl_dwh_recuperaciones, "Tipo_Garantia_Id", 999, "Where A.Tipo_Garantia_Id is null ")
    '3049
    Call cruza_catalogos_2(dbs, tbl_recuperadas_global_vf_inter, tbl_dwh_recuperaciones, "Tipo Cambio", "PROGRAMA", "UDIS", "AGRUPAMIENTO", "TIPO_CREDITO", "TIPO_GARANTIA", "Fecha_Apertura", "SIN FONDOS CONTRAGARANTIA")
    Call campos_calculados_recuperaciones(dbs, tbl_recuperadas_global_vf, tbl_recuperadas_global_vf_inter, "ESTATUS")
    

    
    'Corrige Fecha_Garantia_Honrada de Registro incorrecta
    Call corrige_fecha_pago(dbs, tbl_recuperadas_global_vf)
    Call extracto_bd(dbs, tbl_recuperadas_global_vf, tbl_recuperadas_global_vf & "_sBancomext", " WHERE Producto <> 'GARANTIAS BANCOMEXT' or Producto <>'GARANTIAS SHF/LI FINANCIERO' OR Producto <> 'GARANTIAS BANSEFI'")
    Call extracto_bd(dbs, tbl_recuperadas_global_vf, tbl_recuperadas_global_vf & "_Bancomext", " WHERE (Producto = 'GARANTIAS BANCOMEXT' or Producto ='GARANTIAS SHF/LI FINANCIERO' or Producto is null OR Producto = 'GARANTIAS BANSEFI') ")
    Call cuadro_valida_recup_dwh_dac(dbs, tbl_recuperadas_valida_dwh_dac, tbl_dwh_recuperaciones)
    Call cuadro_valida_recup_td(dbs, tbl_recuperadas_valida_td, tbl_dwh_recuperaciones)
dbs.Close

'Crea Union Flujos
     'Intermediario_Id & Numero_Credito as Concatenar_Saldos
Set dbs = OpenDatabase(db_querie_union_flujos)
    Call agrupa_pagos(dbs, db_querie_pagadas, tbl_pagos_agrup, tbl_pagadas_global_vf)
    Call inserta_columna(dbs, tbl_pagos_agrup, "Concatenar_Saldos", "Text(100)")
    Call corrige_campos(dbs, tbl_pagos_agrup, "Concatenar_Saldos", "A.Intermediario_Id & A.Numero_Credito", "")
    Call importa_recuperaciones(dbs, db_querie_recuperaciones, tbl_recuperadas_global_vf, tbl_recuperadas_global_vf)
    Call agrupa_recuperaciones(dbs, tbl_recuperaciones_agrup, tbl_recuperadas_global_vf)
    Call unir_flujos_p_u_r(dbs, tbl_uf_pagos_recuperaciones, tbl_pagos_agrup, tbl_recuperadas_global_vf)
    
    dbs.Close

    'ACOV 201805 porque la base ya está muy grande
    DAO.DBEngine.CompactDatabase db_querie_union_flujos, wd_processed_dwh & "Querie_UnionFlujos_Temp_" & mes_num & ".accdb"
    Kill db_querie_union_flujos
    'Name Ruta & "Querie_UnionFlujos_Temp_" & mes_num & ".accdb" As db_querie_union_flujos
    Copia = Application.CompactRepair(db_base, wd_processed_dwh & "Copia de Seguridad.accdb")
    Name wd_processed_dwh & "Copia de Seguridad.accdb" As db_querie_union_flujos
    
    Call vincula_tabla(wd_processed_dwh & "Querie_UnionFlujos_Temp_" & mes_num & ".accdb", db_querie_union_flujos, tbl_recuperadas_global_vf, tbl_recuperadas_global_vf)
    Call vincula_tabla(wd_processed_dwh & "Querie_UnionFlujos_Temp_" & mes_num & ".accdb", db_querie_union_flujos, tbl_pagos_agrup, tbl_pagos_agrup)
    Call vincula_tabla(wd_processed_dwh & "Querie_UnionFlujos_Temp_" & mes_num & ".accdb", db_querie_union_flujos, tbl_recuperaciones_agrup, tbl_recuperaciones_agrup)
    
 Set dbs = OpenDatabase(db_querie_union_flujos)
    Call unir_flujos_r_u_p(dbs, tbl_uf_recuperaciones_pagos, tbl_recuperadas_global_vf, tbl_pagos_agrup)
    Call ordena_tabla_final(dbs, tbl_recupera_con_pagos_flujos_ord, tbl_uf_recuperaciones_pagos)
    Call extracto_bd(dbs, tbl_recupera_con_pagos_flujos_ord, tbl_recupera_con_pagos_flujos_ord & "_sBancomext", " WHERE Producto <> 'GARANTIAS BANCOMEXT' or Producto <>'GARANTIAS SHF/LI FINANCIERO' OR Producto <> 'GARANTIAS BANSEFI' ")
    Call extracto_bd(dbs, tbl_recupera_con_pagos_flujos_ord, tbl_recupera_con_pagos_flujos_ord & "_Bancomext", " WHERE (Producto = 'GARANTIAS BANCOMEXT' or Producto ='GARANTIAS SHF/LI FINANCIERO' or Producto is null OR Producto = 'GARANTIAS BANSEFI' ) ")
    Call unir_agrupados_rfc(dbs, tbl_roberto, tbl_pagos_agrup, tbl_recuperaciones_agrup)
dbs.Close
    

'Envio de Bases Finales a base destino y a Excel
'Crea db_querie_pagadas
If existe_ruta(db_pagadas_global_finales) = False Then
    Copia = Application.CompactRepair(db_base_vacia, wd_processed_dwh_bases_finales & "Copia de Seguridad.accdb")
    Name wd_processed_dwh_bases_finales & "Copia de Seguridad.accdb" As db_pagadas_global_finales
End If
'Crea db_querie_recuperaciones
If existe_ruta(db_recupera_con_pagos_flujos_finales) = False Then
    Copia = Application.CompactRepair(db_base_vacia, wd_processed_dwh_bases_finales & "Copia de Seguridad.accdb")
    Name wd_processed_dwh_bases_finales & "Copia de Seguridad.accdb" As db_recupera_con_pagos_flujos_finales
End If

Call copiar_tabla_bd(db_querie_pagadas, db_pagadas_global_finales, tbl_pagadas_global_vf & "_sBancomext", tbl_pagadas_global_vf)
Call copiar_tabla_bd(db_querie_union_flujos, db_recupera_con_pagos_flujos_finales, tbl_recupera_con_pagos_flujos_ord & "_sBancomext", tbl_recupera_con_pagos_flujos)
    
Call exporta_access_excel(db_pagadas_global_finales_xl, db_pagadas_global_finales, tbl_pagadas_global_vf)
Call exporta_access_excel(db_recupera_con_pagos_flujos_finales_xl, db_recupera_con_pagos_flujos_finales, tbl_recupera_con_pagos_flujos)



'Creo las bases de validación de pagos
Set dbs = OpenDatabase(db_querie_pagadas)
dbs.Execute "SELECT [MAX Historico], Producto, Moneda_Id, Sum([SUM Monto_Desembolso]) AS [SumaDeSUM Monto_Desembolso], " _
        & " Sum([SUM Interes_Desembolso]) AS [SumaDeSUM Interes_Desembolso], Sum([SUM Intereses Moratorios]) As [SumaDeSUM Intereses Moratorios] " _
        & " INTO Valida_Pagos FROM [" & tbl_dwh_pagos_f2 & "] Group BY [MAX Historico], Producto, Moneda_Id;"
dbs.Close

Set dbs = OpenDatabase(db_pagadas_global_finales)
dbs.Execute "Select Producto, SUM(Monto_Desembolso_Mn) AS Monto_Desembolso_Mn_Suma, SUM(Interes_Desembolso_Mn) AS Interes_Desembolso_Mn_Suma, " _
        & " SUM(Interes_Moratorios_Mn) AS Interes_Moratorios_Mn_Sum INTO Valida_Base_Pagos_Mn FROM [" & tbl_pagadas_global_vf & "] GROUP BY Producto; "
dbs.Execute "Select Producto, SUM(Monto_Desembolsado) AS Monto_Desembolsado_Suma, SUM(Interes_Desembolso) AS Interes_Desembolso_Suma, " _
        & " SUM(Interes_Moratorios) AS Interes_Moratorios_Sum INTO Valida_Base_Pagos FROM [" & tbl_pagadas_global_vf & "] GROUP BY Producto; "
dbs.Close

Set dbs = OpenDatabase(db_recupera_con_pagos_flujos_finales)
dbs.Execute "Select Producto, SUM(Monto_Mn) AS Monto_Mn_Suma, SUM(Interes_Mn) AS Interes_Mn_Suma, SUM(Moratorios_Mn) AS Moratorios_Mn_Suma, SUM(Excedente_Mn) AS Excedente_Mn_Suma, " _
        & " SUM(Gastos_Juicio_Mn) AS Gastos_Juicio_Mn_Suma INTO Valdia_Base_Pagos_Mn FROM [" & tbl_recupera_con_pagos_flujos & "] GROUP BY Producto;"
dbs.Execute "Select Producto, SUM(Monto) AS Monto_Suma, SUM(Interes) AS Interes_Suma, SUM(Moratorios) AS Moratorios_Suma, SUM(Excedente) AS Excedente_Suma, " _
        & " SUM(Gastos_Juicio_Mn) AS Gastos_Juicio_Mn_Suma INTO Valdia_Base_Pagos FROM [" & tbl_recupera_con_pagos_flujos & "] GROUP BY Producto;"
dbs.Close


'Creo los entregables directamente en la carpeta de Informes Garantías
'Crea Base Pagos
If existe_ruta(db_pagadas_global_entregables) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh_entregables & "Copia de Seguridad.accdb")
    Name wd_processed_dwh_entregables & "Copia de Seguridad.accdb" As db_pagadas_global_entregables
End If

Set dbs = OpenDatabase(db_pagadas_global_entregables)
'Pasa a entregables la base
    dbs.Execute "SELECT * INTO  " & tbl_pagadas_global_vf & "  " _
    & " FROM " & tbl_pagadas_global_vf & " IN '" & db_pagadas_global_finales & "' " _
    & " WHERE PRODUCTO NOT IN ('GARANTIAS BANCOMEXT', 'GARANTIAS BANSEFI', 'GARANTIAS SHF/LI FINANCIERO' ); "
'Base inecesaria
    dbs.Execute "DROP TABLE BD_DWH_Pagos_F1_EmpComp"
'Hace la consulta de validación
    dbs.Execute "Select Producto, SUM(Monto_Desembolso_Mn) AS Monto_Desembolso_Mn_Suma, SUM(Interes_Desembolso_Mn) AS Interes_Desembolso_Mn_Suma, " _
        & " SUM(Interes_Moratorios_Mn) AS Interes_Moratorios_Mn_Sum INTO Valida_Base_Pagos_Mn FROM [" & tbl_pagadas_global_vf & "] GROUP BY Producto; "
    dbs.Execute "Select Producto, SUM(Monto_Desembolsado) AS Monto_Desembolsado_Suma, SUM(Interes_Desembolso) AS Interes_Desembolso_Suma, " _
        & " SUM(Interes_Moratorios) AS Interes_Moratorios_Sum INTO Valida_Base_Pagos FROM [" & tbl_pagadas_global_vf & "] GROUP BY Producto; "

    Call exporta_access_excel(db_pagadas_global_entregables_xl, db_pagadas_global_entregables, tbl_pagadas_global_vf)

dbs.Close


'Lo mimso para recuperaciones
'Crea Base Recup
If existe_ruta(db_recupera_con_pagos_flujos_entregables) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh_entregables & "Copia de Seguridad.accdb")
    Name wd_processed_dwh_entregables & "Copia de Seguridad.accdb" As db_recupera_con_pagos_flujos_entregables
End If

Set dbs = OpenDatabase(db_recupera_con_pagos_flujos_entregables)
'Pasa a entregables la base
    dbs.Execute "SELECT * INTO  " & tbl_recupera_con_pagos_flujos & "  " _
    & " FROM " & tbl_recupera_con_pagos_flujos & " IN '" & db_recupera_con_pagos_flujos_finales & "' " _
    & " WHERE PRODUCTO NOT IN ('GARANTIAS BANCOMEXT', 'GARANTIAS BANSEFI', 'GARANTIAS SHF/LI FINANCIERO' ); "
'Base inecesaria
    dbs.Execute "DROP TABLE BD_DWH_Pagos_F1_EmpComp"
'Hace la consulta de validación
    dbs.Execute "Select Producto, SUM(Monto_Mn) AS Monto_Mn_Suma, SUM(Interes_Mn) AS Interes_Mn_Suma, SUM(Moratorios_Mn) AS Moratorios_Mn_Suma, SUM(Excedente_Mn) AS Excedente_Mn_Suma, " _
        & " SUM(Gastos_Juicio_Mn) AS Gastos_Juicio_Mn_Suma INTO Valdia_Base_Recup_Mn FROM [" & tbl_recupera_con_pagos_flujos & "] GROUP BY Producto;"
    dbs.Execute "Select Producto, SUM(Monto) AS Monto_Suma, SUM(Interes) AS Interes_Suma, SUM(Moratorios) AS Moratorios_Suma, SUM(Excedente) AS Excedente_Suma, " _
        & " SUM(Gastos_Juicio_Mn) AS Gastos_Juicio_Mn_Suma INTO Valdia_Base_Recup FROM [" & tbl_recupera_con_pagos_flujos & "] GROUP BY Producto;"

    Call exporta_access_excel(db_recupera_con_pagos_flujos_entregables_xl, db_recupera_con_pagos_flujos_entregables, tbl_recupera_con_pagos_flujos)

dbs.Close
End Sub




