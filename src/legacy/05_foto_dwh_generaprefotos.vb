Sub GeneraPreFotos()
    Dim dbs As DAO.database
    Dim dbs1 As DAO.database
    Dim Ruta As String
    Dim Engine As DBEngine
    Set Engine = New DBEngine

    Mes = "202412"

    Anio = Left(Mes, 4)
    Mes1 = IIf(Mid(Mes, 5, 1) = 0, Right(Mes, 1), Right(Mes, 2))
    Mes2 = Mes_Letra(Mes1) & Mid(Mes, 3, 2)

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

    db_catalogos = wd_external & "Catálogos_" & Mes2 & ".accdb"
    db_dwh = wd_processed_fotos & "BD_DWH_" & Mes & ".accdb"
    db_dwh_inter = wd_processed_fotos & "BD_DWH_" & Mes & "_Inter.accdb"
    

    tbl_tipo_cambio = "Tipo Cambio"
    tbl_programa = "PROGRAMA"
    tbl_udis = "UDIS"
    tbl_agrupamiento = "AGRUPAMIENTO"
    tbl_tipo_credito = "TIPO_CREDITO"
    tbl_tipo_garantia = "TIPO_GARANTIA"
    tbl_sin_fondos_contragarantia = "SIN FONDOS CONTRAGARANTIA"

    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_tipo_cambio, tbl_tipo_cambio)                     '''Vincula Tipo_Cambio
    'Call Vincula_Tabla(db_catalogos, db_dwh, tbl_tipo_credito, tbl_tipo_credito)             '''Vincula Tipo_Crédito_Id
    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_tipo_garantia, tbl_tipo_garantia)           '''Vincula Tipo_Garantía_Id
    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_udis, tbl_udis)                     '''Vincula Udis
    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_programa, tbl_programa)           '''Vincula Programa
    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_agrupamiento, tbl_agrupamiento)   '''Vincula Agrupamiento
    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_sin_fondos_contragarantia, tbl_sin_fondos_contragarantia) ''Vincula con y sin fondos de contragarantía

    If ExisteRuta(db_dwh_inter) = False Then
        Set dbs1 = Engine.CreateDatabase(db_dwh_inter, dbLangGeneral)
        dbs1.Close
    End If
        For j = 1 To 2
            If j = 1 Then
                Var_NR_R = "R"
                tbl_bd_dwh_nrr_inter = "BD_DWH_NR_" & Mes & "_inter"
                tbl_bd_dwh_nr_r_inter = "BD_DWH_" & Var_NR_R & "_" & Mes & "_inter"
                tbl_bd_dwh_nr_r = "BD_DWH_" & Var_NR_R & "_" & Mes
                tbl_bd_dwh_nr_r_completa = tbl_bd_dwh_nr_r & "_Completa"
            Else
                Var_NR_R = "NR"
                tbl_bd_dwh_nrr_inter = "BD_DWH_R_" & Mes & "_inter"
                tbl_bd_dwh_nr_r_inter = "BD_DWH_" & Var_NR_R & "_" & Mes & "_inter"
                tbl_bd_dwh_nr_r = "BD_DWH_" & Var_NR_R & "_" & Mes
                tbl_bd_dwh_nr_r_completa = tbl_bd_dwh_nr_r & "_Completa"
            End If
            Set dbs = OpenDatabase(db_dwh)
                Inserta_Columna dbs, tbl_bd_dwh_nr_r, "Fecha_Consulta", "Date"
                Corrige_Campos dbs, tbl_bd_dwh_nr_r, "Fecha_Consulta", "#" & Mes1 & "/01/" & Anio & "#", ""
                
                Corrige_Campos dbs, tbl_bd_dwh_nr_r, "Tipo_Garantia_Id", 999, "Where A.Tipo_Garantia_Id is null "
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Call Cruza_Catalogos(dbs, tbl_bd_dwh_nr_r_completa, tbl_bd_dwh_nr_r, tbl_tipo_cambio, tbl_programa, tbl_udis, tbl_agrupamiento, tbl_tipo_credito, tbl_tipo_garantia, tbl_sin_fondos_contragarantia)
                'Call Cruza_Catalogos(dbs, tbl_bd_dwh_nr_r_completa, tbl_bd_dwh_nr_r, tbl_tipo_cambio, tbl_programa, tbl_udis, tbl_agrupamiento, tbl_tipo_credito, tbl_tipo_garantia)
            dbs.Close
            
            'compacta_repara (db_dwh)
            Call Vincula_Tabla(db_dwh, db_dwh_inter, tbl_bd_dwh_nr_r_completa, tbl_bd_dwh_nr_r_completa)

        Next j

End Sub