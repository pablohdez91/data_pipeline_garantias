Sub SubeSaldosyMGI()
Dim dbs As database

Mes = "202411"

wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
wd_external = wd & "data\external\"
wd_processed = wd & "data\processed\"
wd_processed_dwh = wd_processed & "DWH\"
wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
wd_processed_fotos = wd_processed & "Fotos\"
wd_raw = wd & "data\raw\"
wd_staging = wd & "data\staging\"
wd_validations = wd & "data\validations\"

tbl_datos = "DATOS"
db_saldo_mgi = wd_raw & "Saldo_MGI_" & Mes & ".mdb"
db_saldo_mgi_80686 = wd_raw & "Saldo_MGI_80686_" & Mes & ".mdb" 'ACOV201812 Jera quiere una sola tabla de saldos
db_dwh_inter = wd_processed_fotos & "BD_DWH_" & Mes & "_Inter.accdb"

tbl_bd_dwh_saldoymgi = "BD_DWH_SaldoyMGI_" & Mes
tbl_bd_dwh_saldoymgi_80686 = "BD_DWH_SaldoyMGI_80686_" & Mes

Call Vincula_Tabla(db_saldo_mgi, db_dwh_inter, tbl_datos, tbl_bd_dwh_saldoymgi)
Call Vincula_Tabla(db_saldo_mgi_80686, db_dwh_inter, tbl_datos, tbl_bd_dwh_saldoymgi_80686)

Set dbs = OpenDatabase(db_dwh_inter)
    cadena = " INTERMEDIARIO_ID, MONEDA_ID, NUMERO_CREDITO, [MONTO_GARANTIZADO (SUMA)] as Monto_Garantizado, [SALDO_CONTINGENTE (SUMA)] as Saldo_Contingente "
    
    'ACOV 201812 para unir las tablas de saldo
    dbs.Execute "SELECT * INTO AUX FROM " & tbl_bd_dwh_saldoymgi & " ; "
    dbs.Execute "INSERT INTO AUX SELECT * FROM " & tbl_bd_dwh_saldoymgi_80686 & ";"
    
    'Call Crea_Tabla_Agrega_Campos(dbs, cadena, ", ", tbl_bd_dwh_saldoymgi, tbl_bd_dwh_saldoymgi & "_VF", " A.Intermediario_Id & A.Numero_Credito as Concatenado ", "", "", "")
    Call Crea_Tabla_Agrega_Campos(dbs, cadena, ", ", "AUX", tbl_bd_dwh_saldoymgi & "_VF", "A.Intermediario_Id & A.Numero_Credito as Concatenado ", "", "", "")
dbs.Close

End Sub
