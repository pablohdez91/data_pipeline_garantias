Sub Principal()
Dim dbs As DAO.Database

MesNum = 202411

'Datos Calculados
Anio = Left(MesNum, 4)
Mes1 = IIf(Mid(MesNum, 5, 1) = 0, Right(MesNum, 1), Right(MesNum, 2))
mes = Mes_Letra(Mes1) & Mid(MesNum, 3, 2)
Mes_LNTG = Mes_Letra_LNTG(Mes1)
Carpeta_LNTG = Mes_LNTG & Anio

'Carpetas
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

'Inputs
db_base = wd_staging & "Base.accdb"
db_catalogos = wd_external & "Catálogos_" & mes & ".accdb"
db_bd_dwh_vf = wd_processed_fotos & "BD_DWH_" & MesNum & "_VF.accdb"

'Outputs
db_querie_saldos = wd_processed_dwh & "Querie_Saldos_" & MesNum & ".accdb"
db_prem = wd_raw & "Prem_" & MesNum & ".mdb"
db_dwh_plazo_rem = wd_raw & "BD_DWH_PlazoRem_" & MesNum & ".xlsx"


tbl_base_agrup_final = "BASE_AGRUP_FINAL_" & MesNum
tbl_estados = "ESTADOS"
tbl_sector = "SECTOR"
tbl_estrato = "ESTRATO"
tbl_tipo_garantia = "TIPO_GARANTIA"
tbl_tipo_credito = "TIPO_CREDITO"
tbl_sin_fondos_contragarantia = "SIN FONDOS CONTRAGARANTIA"

agrup_1 = "BUCKET, Producto, Tipo_Credito_Id, MM_UDIS, TPRO_CLAVE, nr_r, CSG, Intermediario_Id, [Razón Social (Intermediario)], TIPO_PERSONA, [Porcentaje de Comisión Garantia], [Porcentaje Garantizado], [AGRUPAMIENTO], Estado, Sector, Estrato, [Tipo_garantia], [Tipo_Credito], Programa_Id, CSF"
AGRUP = "BUCKET, Producto, Tipo_Credito_Id, MM_UDIS, TPRO_CLAVE, nr_r, CSG, Intermediario_Id, [Razón Social (Intermediario)], TIPO_PERSONA, [Porcentaje de Comisión Garantia], [Porcentaje Garantizado], [AGRUPAMIENTO], Estado, Sector, Estrato, [Tipo_garantia], [Tipo_Credito], Programa_Id,  CSF"

'Crea Base_Saldos_Agrupados
If ExisteRuta(db_querie_saldos) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_dwh & "Copia de Seguridad.accdb")
    Name wd_processed_dwh & "Copia de Seguridad.accdb" As db_querie_saldos
End If

Set dbs = OpenDatabase(db_querie_saldos)
    Call Vincula_Tabla(db_catalogos, db_querie_saldos, tbl_estados, tbl_estados)             '''Vincula Estado
    Call Vincula_Tabla(db_catalogos, db_querie_saldos, tbl_sector, tbl_sector)             '''Vincula Sector
    Call Vincula_Tabla(db_catalogos, db_querie_saldos, tbl_estrato, tbl_estrato)           '''Vincula Estrato
    Call Vincula_Tabla(db_catalogos, db_querie_saldos, tbl_tipo_garantia, tbl_tipo_garantia)
    Call Vincula_Tabla(db_catalogos, db_querie_saldos, tbl_tipo_credito, tbl_tipo_credito)
    Call Vincula_Tabla(db_catalogos, db_querie_saldos, tbl_sin_fondos_contragarantia, tbl_sin_fondos_contragarantia)
    For k = 1 To 2
        If k = 1 Then
            nr_r = "NR"
        Else
            nr_r = "R"
        End If
        tbl_bd_dwh_nrr = "BD_DWH_" & nr_r & "_" & MesNum
        Call Garantias_c_Saldo(dbs, db_bd_dwh_vf, tbl_bd_dwh_nrr, tbl_bd_dwh_nrr)
        Call Cruza_Catalogos_2(dbs, tbl_bd_dwh_nrr & "_Completo", tbl_bd_dwh_nrr, tbl_estados, tbl_estrato, tbl_sector, tbl_tipo_garantia, tbl_tipo_credito, tbl_sin_fondos_contragarantia)
        Call Agrupa_BDcSaldo(dbs, tbl_bd_dwh_nrr & "_VF", tbl_bd_dwh_nrr & "_Completo", AGRUP, agrup_1)
        If k = 1 Then
            Call Crea_Tabla(dbs, "A.* ", tbl_bd_dwh_nrr & "_VF", tbl_base_agrup_final, "", "", "")
        Else
            Call Inserta_Filas(dbs, "A.* ", tbl_bd_dwh_nrr & "_VF", tbl_base_agrup_final, "", "", "")
        End If
    Next k
dbs.Close

Set dbs = OpenDatabase(db_prem)
    dbs.Execute " SELECT BUCKET, DESC_INDICADOR AS Producto, FECHA_APERTURA as [Fecha de Apertura], FECHA_CONSULTA, FVTO_RIESGOSD, " _
        & " INDICADOR_ID as [Producto ID], INTERMEDIARIO_ID, ORDEN_FREGISTRO, PROGRAMA_ID, RAZON_SOCIAL as [Razón Social (Intermediario)], " _
        & " [MONTO_CREDITO_MN (SUMA)] as [SUM Monto _Credito_Mn], [NUMERO_CREDITOS (SUMA)] as [COUNT Numero_Credito], " _
        & " [SALDO_CONTINGENTE_MN (SUMA)] as  [SUM Saldo_Contingente_Mn] INTO BD_DWH_PlazoRem_" & MesNum & " FROM DATOS;"
        
    dbs.Execute "SELECT Producto, Sum([SUM Saldo_Contingente_Mn]) AS S_Saldo INTO Valida_Saldo FROM BD_DWH_PlazoRem_" & MesNum & " GROUP BY Producto;"
    
    Call Exporta_Access_a_Excel(db_dwh_plazo_rem, db_prem, "BD_DWH_PlazoRem_" & MesNum)
    
dbs.Close
End Sub



Function Garantias_c_Saldo(dbs, BaseOrigen, Tabla_I, Tabla_F)
dbs.Execute "SELECT * " _
          & "INTO " & Tabla_F & " " _
          & "FROM " & Tabla_I & " IN '" & BaseOrigen & "' " _
          & "WHERE Saldo_contingente_Mn > 0;"
End Function
Function Cruza_Catalogos_2(dbs, T_Final, T_Iiquierda, T_Estado, T_Estrato, T_Sector, T_TipoGarantia, T_TipoCredito, T_Sinfondos)
    dbs.Execute "SELECT A.*, B.Estado, C.Estrato, D.Sector, E.Tipo_garantia, F.Tipo_Credito , IIF(G.FONDOS_CONTRAGARANTIA='SF','SF','CF') as CSF " _
        & "INTO [" & T_Final & "] " _
        & "FROM ((((([" & T_Iiquierda & "] as A LEFT JOIN [" & T_Estado & "] as B " _
        & "ON A.[Estado_Id]=B.[Estado_ID]) LEFT JOIN [" & T_Estrato & "] as C " _
        & "ON A.[Estrato_Id]=C.[Estrato_Id])  LEFT JOIN [" & T_Sector & "] as D " _
        & "ON A.[Sector_Id]=D.[Sector_ID])  LEFT JOIN [" & T_TipoGarantia & "] as E " _
        & "ON A.[Tipo_Garantia_Id]=E.[Tipo_garantia_ID])  LEFT JOIN [" & T_TipoCredito & "] as F " _
        & "ON A.[Tipo_Credito_Id]=F.[Tipo_Credito_ID]) LEFT JOIN [" & T_Sinfondos & "] as G " _
        & "ON (cstr(A.[Intermediario_Id])=G.[Intermediario_Id] AND A.[Numero_Credito]=G.[CLAVE_CREDITO]);"
End Function
Function Agrupa_BDcSaldo(dbs, T_Final, T_Inicial, AGRUP, AGRUP1)
dbs.Execute "SELECT " & AGRUP1 & ", " _
          & "SUM([Saldo_Contingente_Mn]) AS SALDO_MN, COUNT(Numero_Credito) AS TOT_CREDITOS " _
          & "INTO " & T_Final & " " _
          & "FROM " & T_Inicial & " " _
          & "GROUP BY " & AGRUP & "; "
End Function
Function Crea_Tabla(dbs, Linea_1, TablaInicial, TablaFinal, BaseDestino, Filtro, Agrupado_por)
    dbs.Execute "select " & Linea_1 & " " _
        & "into [" & TablaFinal & "] " _
        & "from " & TablaInicial & " as A " _
        & " " & BaseDestino & " " _
        & " " & Filtro & " " _
        & " " & Agrupado_por & "; "
End Function
Function Inserta_Filas(dbs, Linea_1, TablaInicial, TablaFinal, BaseDestino, Filtro, Agrupado_por)
    dbs.Execute "insert into [" & TablaFinal & "] " _
        & "select " & Linea_1 & " " _
        & "from [" & TablaInicial & "] as A " _
        & " " & BaseDestino & " " _
        & " " & Filtro & " " _
        & " " & Agrupado_por & "; "
End Function