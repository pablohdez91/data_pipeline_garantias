Sub Principal()
Dim dbs As DAO.Database

MesNum = 202408


'ACOV 201808 Fide dejó de tener saldo en microcrédito entonces realmente no se neceista ya integrarlo a la base de los agrup

'Datos Calculados
Anio = Left(MesNum, 4)
Mes1 = IIf(Mid(MesNum, 5, 1) = 0, Right(MesNum, 1), Right(MesNum, 2))
mes = Mes_Letra(Mes1) & Mid(MesNum, 3, 2)
Mes_LNTG = Mes_Letra_LNTG(Mes1)
Carpeta_LNTG = Mes_LNTG & Anio

'BD_DWH = "Saldos_FIDE_" & MesNum
'BaseTBL = wd_raw & BD_DWH & ".mdb"

wd = "D:\DAR\proyecto_mejora_fotos\2. Nuevas fotos\"
wd_external = wd & "data\external\"
wd_processed = wd & "data\processed\"
wd_processed_dwh = wd_processed & "DWH\"
wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
wd_processed_dwh_bases_finales_saldos = wd_processed_dwh_bases_finales & "saldos\"
wd_processed_fotos = wd_processed & "Fotos\"
wd_raw = wd & "data\raw\"
wd_staging = wd & "data\staging\"
wd_validations = wd & "data\validations\"

db_catalogos = wd_external & "Catálogos_" & mes & ".accdb"
db_base = wd_staging & "Base.accdb"

db_querie_saldos_integrados = wd_processed_dwh_bases_finales_saldos & "Querie_SaldosIntegrados_" & MesNum & ".accdb"
db_qurie_saldos = wd_processed_dwh & "Querie_Saldos_" & MesNum & ".accdb"
db_base_agrup_final = wd_processed_dwh_bases_finales_saldos & "BASE_AGRUP_FINAL_" & MesNum & ".accdb"

tbl_base_agrup_final_fide = "BASE_AGRUP_FINAL_" & MesNum & "_FIDE"
tbl_base_agrup_final = "BASE_AGRUP_FINAL_" & MesNum

If ExisteRuta(wd_processed_dwh_bases_finales_saldos) = False Then
    MkDir (wd_processed_dwh_bases_finales_saldos)
End If

'Crea Base_Saldos_Agrupados
If ExisteRuta(db_querie_saldos_integrados) = False Then
    Copia = Application.CompactRepair(db_base, wd_raw & "Copia de Seguridad.accdb")
    Name wd_raw & "Copia de Seguridad.accdb" As db_querie_saldos_integrados
End If

tbl_estados = "ESTADOS"
tbl_sector = "SECTOR"
tbl_estrato = "ESTRATO"
tbl_tipo_garantia = "TIPO_GARANTIA"
tbl_tipo_credito = "TIPO_CREDITO"

Set dbs = OpenDatabase(db_querie_saldos_integrados)
    'Call Vincula_Tabla(BaseTBL, db_querie_saldos_integrados, "DATOS", "AUX_FIDE_TBL")   'ACOV 201808 ya no tiene saldo enotnces no existe la consulta ACOV 201805 a raíz del cambio a tableau
    Call Vincula_Tabla(db_catalogos, db_querie_saldos_integrados, tbl_estados, tbl_estados)             '''Vincula Estado
    Call Vincula_Tabla(db_catalogos, db_querie_saldos_integrados, tbl_sector, tbl_sector)             '''Vincula Sector
    Call Vincula_Tabla(db_catalogos, db_querie_saldos_integrados, tbl_estrato, tbl_estrato)           '''Vincula Estrato
    Call Vincula_Tabla(db_catalogos, db_querie_saldos_integrados, tbl_tipo_garantia, tbl_tipo_garantia)
    Call Vincula_Tabla(db_catalogos, db_querie_saldos_integrados, tbl_tipo_credito, tbl_tipo_credito)
    
    'ACOV 201808 ya no tiene saldo enotnces no existe la consulta
    'ACOV 201805 proceso para actualizar los nombres
    'dbs.Execute "SELECT ESTADO_ID, ESTRATO_DESCRIPCION as Estrato, ESTRATO_ID, INTERMEDIARIO_ID, PORCENTAJE_COMISION_GARANTIA, PORCENTAJE_GARANTIZADO, " _
        & " PROGRAMA_ID, PROGRAMA_ORIGINAL, SECTOR_ID, TIPO_CREDITO_DESCRIPCION, TIPO_CREDITO_ID, TIPO_GARANTIA_ID, TPRO_CLAVE, " _
        & " [SALDO_CONTINGENTE_MN (SUMA)] as [SUM Saldo_Contingente_Mn] INTO " & BD_DWH & " FROM AUX_FIDE_TBL; "
    
    'Sigue proceso anterior
    'Call Cruza_Catalogos_2(dbs, BD_DWH & "_Completo", BD_DWH, tbl_estados, tbl_estrato, tbl_sector, tbl_tipo_garantia, tbl_tipo_credito)
    'Call IntegraSaldosFIDE(dbs, tbl_base_agrup_final_fide, BD_DWH & "_Completo")
    'Call Corrige_Campos(dbs, tbl_base_agrup_final_fide, "TPRO_CLAVE", "IIf(A.Programa_Id>=32000 And A.Programa_Id<=32100, A.Programa_Id, IIf(A.Programa_Id=3976 And A.Progama_Original=31415,A.Programa_Id,IIf(A.Progama_Original = 3999,A.Programa_Id,A.Progama_Original))) ", "")  'Falata ProgramaOriginal
    Call Crea_Tabla(dbs, tbl_base_agrup_final, tbl_base_agrup_final, " IN '" & db_qurie_saldos & "' ", "", "")
    'Call Inserta_Filas(dbs, "A.*,'CF' as CSF ", tbl_base_agrup_final_fide, tbl_base_agrup_final, "", "", "")
dbs.Close

'Tabla de validación de los saldos ACOV201805
Set dbs = OpenDatabase(db_querie_saldos_integrados)
dbs.Execute "SELECT Producto, SUM(SALDO_MN) AS [S_SALDO(MDP)] INTO VALIDA_SALDO FROM BASE_AGRUP_FINAL_" & MesNum & " GROUP BY Producto"

'Importa base a Excel ACOV201805
db_base_agrup_final_xl = wd_processed_dwh_bases_finales_saldos & "BASE_AGRUP_FINAL_" & MesNum & ".xlsx"

Call Exporta_Access_a_Excel(db_base_agrup_final_xl, db_querie_saldos_integrados, tbl_base_agrup_final)
Copia = Application.CompactRepair(db_base, wd_raw & "Copia de Seguridad.accdb")
Name wd_raw & "Copia de Seguridad.accdb" As db_base_agrup_final
Set dbs = OpenDatabase(db_base_agrup_final)
    dbs.Execute "SELECT * INTO BASE_AGRUP_FINAL_" & MesNum & " FROM BASE_AGRUP_FINAL_" & MesNum & " IN '" & db_querie_saldos_integrados & "' ;"
    'dbs.Execute "DROP TABLE BD_DWH_Pagos_F1_EmpComp"
dbs.Close
End Sub



Function Crea_Tabla(dbs, TablaInicial, TablaFinal, BaseDestino, Filtro, Agrupado_por)
    dbs.Execute "select A.BUCKET, A.Producto, A.TPRO_CLAVE, A.NR_R, " _
        & "A.CSG, A.Intermediario_Id, A.[Razón Social (Intermediario)], A.TIPO_PERSONA, " _
        & "A.[Porcentaje de Comisión Garantia], A.[Porcentaje Garantizado], " _
        & "A.AGRUPAMIENTO, A.Estado, A.Sector, A.Estrato, A.[Tipo_garantia], A.[Tipo_Credito], A.Programa_Id, A.SALDO_MN, A.CSF " _
        & "into [" & TablaFinal & "] " _
        & "from " & TablaInicial & " as A " _
        & " " & BaseDestino & " " _
        & " " & Filtro & " " _
        & " " & Agrupado_por & "; "
End Function
