Private Sub Principal()

Dim dbs As DAO.Database
MesNum = 202410
Anio = Left(MesNum, 4)
Mes1 = IIf(Mid(MesNum, 5, 1) = 0, Right(MesNum, 1), Right(MesNum, 2))
mes = mes_letra(Mes1) & Mid(MesNum, 3, 2)

' Prefixes
'   wd: Working Directory
'   db: Database (Access file)
'   tbl: Table

wd = "D:\DAR\proyecto_mejora_fotos\2. Nuevas fotos\"
wd_external = wd & "data\external\"
wd_processed = wd & "data\processed\"
wd_processed_dwh = wd_processed & "DWH\"
'wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
'wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
wd_raw = wd & "data\raw\"
wd_staging = wd & "data\staging\"
'wd_validations = wd & "data\validations\"

'BDs Inputs
db_catalogos = wd_external & "Catálogos_" & mes & ".accdb"
db_base = wd_staging & "Base.accdb"
db_saldos = wd_raw & "Saldos_" & MesNum & ".mdb"
db_saldos_bmxt = wd_raw & "Saldos_BMXT_" & MesNum & ".mdb"
db_saldos_80686 = wd_raw & "Saldos_80686_" & MesNum & ".mdb" 'ACOV 201812 Jera no necesita las bases seprardas

'BDs Outputs
db_dwh_saldos_totales = wd_processed_dwh & "BD_DWH_SaldosTotales_" & MesNum & ".accdb"

' Tablas
tbl_dwh_saldos_totales = "DWH_SaldosTotales_" & MesNum
tbl_dwh_agrup_saldo_total = "BD_DWH_Agrup_SaldoTotal_" & MesNum
tbl_dwh_agruptaxo_saldo_total = "BD_DWH_AgrupTaxo_SaldoTotal_" & MesNum
tbl_tipo_credito = "TIPO_CREDITO"

agrupamiento = "Producto, [Razón Social (Intermediario)], Intermediario_Id, Programa_Original, Programa_Id"
agrupamiento_taxtonomia = "Producto"

'Crea Base_Saldos_Totales
If existe_ruta(db_dwh_saldos_totales) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh & "Copia de Seguridad.accdb")
    Name wd_processed_dwh & "Copia de Seguridad.accdb" As db_dwh_saldos_totales
End If

'ACOV 201805 Vincula las bases de saldo
Call vincula_tabla(db_saldos, db_dwh_saldos_totales, "DATOS", "AUX1")
Call vincula_tabla(db_saldos_bmxt, db_dwh_saldos_totales, "DATOS", "AUX2")
Call vincula_tabla(db_saldos_80686, db_dwh_saldos_totales, "DATOS", "AUX3")

Set dbs = OpenDatabase(db_dwh_saldos_totales)
    
    'ACOV 201805 cambio por la migración a TBL renombro las columnas
    dbs.Execute "SELECT * INTO " & tbl_dwh_saldos_totales & "_AUX FROM AUX1;"
    dbs.Execute "INSERT INTO " & tbl_dwh_saldos_totales & "_AUX SELECT * FROM AUX2 ; "
    dbs.Execute "INSERT INTO " & tbl_dwh_saldos_totales & "_AUX SELECT * FROM AUX3; "
    
    dbs.Execute "SELECT DESC_INDICADOR AS Producto, FECHA_CONSULTA, INDICADOR_ID AS [Producto ID], INTERMEDIARIO_ID, " _
        & " PROGRAMA_DESCRIPCION AS [Programa Descripción (ID)], PROGRAMA_ID, PROGRAMA_ORIGINAL,  RAZON_SOCIAL AS [Razón Social (Intermediario)], " _
        & " TIPO_CREDITO_ID, [SALDO_CONTINGENTE_MN (SUMA)] AS [SUM Saldo_Contingente_Mn] INTO " & tbl_dwh_saldos_totales & " FROM " & tbl_dwh_saldos_totales & "_AUX ; "
    dbs.Execute "DROP TABLE " & tbl_dwh_saldos_totales & "_AUX; "
    
    'Agrega columna TPRO_CLAVE
    'ACOV 201805 se agregó las de santander y empresa mediana (se corrigió el progama por programa)
    Call inserta_columna(dbs, tbl_dwh_saldos_totales, "TPRO_CLAVE", "Double")
    Call corrige_campos(dbs, tbl_dwh_saldos_totales, "TPRO_CLAVE", "IIf(A.Programa_Id>=32000 And A.Programa_Id<=32100, A.Programa_Id, IIf(A.Programa_Id=3976 And A.Programa_Original=31415,A.Programa_Id,IIF(A.Programa_Original = 33842 AND A.Programa_Id = 33366, A.Programa_Id, IIF(A.Programa_Original = 3200 AND A.Programa_Id IN (3536, 3537, 3539, 3542,3544, 3545, 3546,3547,3548,3549,3550, 3553, 3555, 3558,3559, 3560, 3564,3566), IIf(A.Programa_Original = 3999,A.Programa_Id,A.Programa_Original))))) ", "")
    Call vincula_tabla(db_catalogos, db_dwh_saldos_totales, tbl_tipo_credito, tbl_tipo_credito)
    Call cruza_catalogos(dbs, tbl_dwh_saldos_totales & "_VF", tbl_dwh_saldos_totales, tbl_tipo_credito)
    
    filtro_1 = " HAVING ((Producto = 'GARANTIAS BANCOMEXT') or (Producto = 'GARANTIAS SHF/LI FINANCIERO') or (Producto = 'GARANTIAS BANSEFI') or ([Razón Social (Intermediario)] = 'FISO PARA EL AHORRO DE ENERGIA ELECTRICA (FASE II)')) "
    filtro_2 = " HAVING Producto <> 'GARANTIAS BANCOMEXT'and Producto <> 'GARANTIAS SHF/LI FINANCIERO' and Producto <> 'GARANTIAS BANSEFI' and [Razón Social (Intermediario)] <>'FISO PARA EL AHORRO DE ENERGIA ELECTRICA (FASE II)' "
    filtro_taxonomia_1 = " HAVING ((Producto = 'GARANTIAS BANCOMEXT') or (Producto = 'GARANTIAS SHF/LI FINANCIERO')) or (Producto = 'GARANTIAS BANSEFI') "
    filtro_taxonomia_2 = " HAVING Producto <> 'GARANTIAS BANCOMEXT'and Producto <> 'GARANTIAS SHF/LI FINANCIERO' and Producto <> 'GARANTIAS BANSEFI' "
    
    Call agrupa_db_saldos_totales(dbs, tbl_dwh_agrup_saldo_total, tbl_dwh_saldos_totales & "_VF", " ", agrupamiento)
    Call agrupa_db_saldos_totales(dbs, tbl_dwh_agrup_saldo_total & "_Resto", tbl_dwh_saldos_totales & "_VF", filtro_1, agrupamiento)
    Call agrupa_db_saldos_totales(dbs, tbl_dwh_agrup_saldo_total & "_Automatica", tbl_dwh_saldos_totales & "_VF", filtro_2, agrupamiento)
    Call agrupa_db_saldos_totales(dbs, tbl_dwh_agruptaxo_saldo_total, tbl_dwh_saldos_totales & "_VF", " ", agrupamiento_taxtonomia)
    Call agrupa_db_saldos_totales(dbs, tbl_dwh_agruptaxo_saldo_total & "_Automatica", tbl_dwh_saldos_totales & "_VF", filtro_taxonomia_2, agrupamiento_taxtonomia)
    Call agrupa_db_saldos_totales(dbs, tbl_dwh_agruptaxo_saldo_total & "_Resto", tbl_dwh_saldos_totales & "_VF", filtro_taxonomia_1, agrupamiento_taxtonomia)
dbs.Close
End Sub




Function agrupa_db_saldos_totales(dbs, T_Final, T_Inicial, Filtro, agrupamiento)
dbs.Execute "SELECT " & agrupamiento & ", " _
          & "SUM([SUM Saldo_Contingente_Mn]) AS SALDO_MN " _
          & "INTO [" & T_Final & "] " _
          & "FROM [" & T_Inicial & "] " _
          & "GROUP BY " & agrupamiento & " " _
          & " " & Filtro & " ; "
End Function

Function corrige_campos(dbs, TablaMod, Columna, Condicion, Filtro)
    dbs.Execute "update [" & TablaMod & "] as A " _
    & "set [" & Columna & "] = " & Condicion & " " _
    & " " & Filtro & ";"
End Function

Function cruza_catalogos(dbs, T_Final, T_Izquierda, tbl_tipo_credito)
    dbs.Execute "SELECT A.*, B.NR_R " _
        & "INTO [" & T_Final & "] " _
        & "FROM [" & T_Izquierda & "] as A LEFT JOIN [" & tbl_tipo_credito & "] as B " _
        & "ON (A.[Tipo_Credito_Id]=B.[Tipo_Credito_ID])" _
        & "Where Producto <> null ; "
End Function

Function existe_ruta(ByVal Ruta As String) As Boolean
    existe_ruta = False
    If Dir(Ruta, vbDirectory) = "" Then existe_ruta = False _
    Else existe_ruta = True
    If existe_ruta = False Then
        On Local Error Resume Next
         existe_ruta = Len(Dir$(Ruta))
        If Err Then
            existe_ruta = False
        End If
        Err = 0
        On Local Error GoTo 0
    End If
End Function

Function inserta_columna(dbs, TablaMod, Columna, TP_Columna)
    dbs.Execute "alter table [" & TablaMod & "] " _
    & "add column [" & Columna & "] " & TP_Columna & " ; "
End Function


Function mes_letra(mes) As String
    Select Case mes
    Case 1
        mes_letra = "Enero"
    Case 2
        mes_letra = "Febrero"
    Case 3
        mes_letra = "Marzo"
    Case 4
        mes_letra = "Abril"
    Case 5
        mes_letra = "Mayo"
    Case 6
        mes_letra = "Junio"
    Case 7
        mes_letra = "Julio"
    Case 8
        mes_letra = "Agosto"
    Case 9
        mes_letra = "Septiembre"
    Case 10
        mes_letra = "Octubre"
    Case 11
        mes_letra = "Noviembre"
    Case 12
        mes_letra = "Diciembre"
    End Select
    mes_letra = Left(mes_letra, 3)
End Function

Function vincula_tabla(BaseOrigen, BaseDestino, tbl_dwh_saldos_totales, TablaDestino)
    Dim objetAccessO As Access.Application
    Set objetAccessO = New Access.Application
    objetAccessO.OpenCurrentDatabase BaseDestino
    objetAccessO.DoCmd.TransferDatabase acLink, "Microsoft Access", BaseOrigen, acTable, tbl_dwh_saldos_totales, TablaDestino
    objetAccessO.Quit
    Set objetAccessO = Nothing
End Function

