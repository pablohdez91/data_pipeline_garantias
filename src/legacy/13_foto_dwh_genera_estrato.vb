Sub BuscaMaxFechaApertura()
    Mes = "202412"

    wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_dwh = wd_processed & "DWH\"
    wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
    wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
    wd_processed_fotos = wd_processed & "Fotos\"
    wd_processed_fotos_cierre = wd_processed_fotos & Mes & "\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"

    db_foto_revolventes_cohortes_preliminar = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & Mes & "_Preeliminar.accdb"


    Linea = "A.NOMBRE, A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.TAXONOMIA, A.INTER_CLAVE, A.NR_R, A.CSG "
    tbl_vf_foto_r = "VF_Foto_R_" & Mes
    db_origen = ""
    Dim dbs As database
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "SELECT " _
        & "" & Linea & ", MAX(A.FECHA_VALOR) AS FECHA_VALOR23 " _
        & "INTO [Temp_Estrato] " _
        & "FROM [" & tbl_vf_foto_r & "] as A " _
        & "IN '" & db_origen & "'" _
        & "GROUP BY  " & Linea & " " _
        & "ORDER BY " & Linea & " ; "
        
        Call Crea_Tabla_Agrega_Campos_Cruzada(dbs, Linea, ", Max(A.FECHA_VALOR23) as Max_Fecha_Valor, Max(B.Estrato_Id) as Estrato_Id", "Temp_Estrato", tbl_vf_foto_r, "VF_Estrato_" & Mes, "", "", " ON (A.NOMBRE=B.NOMBRE AND A.BANCO=B.BANCO AND A.AGRUPAMIENTO=B.AGRUPAMIENTO AND A.AGRUPAMIENTO_ID=B.AGRUPAMIENTO_ID AND A.TAXONOMIA=B.TAXONOMIA AND A.INTER_CLAVE=B.INTER_CLAVE AND A.NR_R=B.NR_R AND A.CSG=B.CSG AND A.FECHA_VALOR23 = B.FECHA_VALOR)", "GROUP BY  " & Linea)
        Call Borrar_Tabla(dbs, "Temp_Estrato")
    dbs.Close
End Sub


Function Crea_Tabla_Agrega_Campos_Cruzada(dbs, linea_1, Coma, TablaInicial_1, TablaInicial_2, TablaFinal, Campos_Extra, BaseOrigen, Filtro, Agrupado_por)
    dbs.Execute "select " & linea_1 & Coma & " " _
        & "" & Campos_Extra & " " _
        & "into [" & TablaFinal & "] " _
        & "from [" & TablaInicial_1 & "] as A left join [" & TablaInicial_2 & "] as B " _
        & "" & BaseOrigen & " " _
        & "" & Filtro & " " _
        & "" & Agrupado_por & ";"
End Function



