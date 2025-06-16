Sub BuscaMaxFechaApertura() '(BaseDestino, db_foto_revolventes_cohortes_preliminar, tbl_vf_foto_r, linea)
    Dim Engine As DBEngine
    Dim dbs As database
    Set Engine = New DBEngine
    Dim tdf As TableDef
    Dim n As Object

    Mes = "202411"
    
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
    db_foto_revolventes_cohortes = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & Mes & "_VF.accdb"
    db_foto_simples_cohortes = wd_processed_fotos_cierre & "FotoSimplesCohortes_" & Mes & "_VF.accdb"

    Linea = "A.NOMBRE, A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.TAXONOMIA, A.INTER_CLAVE, A.NR_R, A.CSG "
    tbl_vf_foto_r = "VF_Foto_R_" & Mes


    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    
    'Call Elimina_Columna(dbs, TablaMod, Columna)
    tbl_estrato_id = "Estrato_Id"
    tbl_estrato_id_original = "Estrato_Id_Original"
    tbl_tipo_persona = "TIPO_PERSONA"
    tbl_tipo_persona_original = "TIPO_PERSONA_Original"


    Set tdf = dbs.TableDefs(tbl_vf_foto_r)
    For Each n In tdf.Fields
        If n.Name = tbl_estrato_id Then n.Name = tbl_estrato_id_original
        If n.Name = tbl_tipo_persona Then n.Name = tbl_tipo_persona_original
        'End If
    Next n
    Set tdf = Nothing
    dbs.Close
        
    Set dbs = Engine.CreateDatabase(db_foto_revolventes_cohortes, dbLangGeneral)
    dbs.Close
     
    Vincula_Tabla db_foto_revolventes_cohortes_preliminar, db_foto_revolventes_cohortes, "VF_PFPM_" & Mes, "VF_PFPM_" & Mes
    
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes)
    Call Crea_Tabla_Agrega_Campos_Cruzada(dbs, "A.*", ", B.Estrato_Id as Estrato_Id ", tbl_vf_foto_r, "VF_Estrato_" & Mes, "Temp_" & tbl_vf_foto_r, "", " ON (A.NOMBRE=B.NOMBRE AND A.BANCO=B.BANCO AND A.AGRUPAMIENTO=B.AGRUPAMIENTO AND A.AGRUPAMIENTO_ID=B.AGRUPAMIENTO_ID AND A.TAXONOMIA=B.TAXONOMIA AND A.INTER_CLAVE=B.INTER_CLAVE AND A.NR_R=B.NR_R AND A.CSG=B.CSG)", " IN '" & db_foto_revolventes_cohortes_preliminar & "' ", "")
    Call Crea_Tabla_Agrega_Campos_Cruzada(dbs, "A.*", ", B.TIPO_PERSONA as TIPO_PERSONA ", "Temp_" & tbl_vf_foto_r, "VF_PFPM_" & Mes, tbl_vf_foto_r & "_VCohortes", "", "", " ON (A.NOMBRE=B.NOMBRE AND A.BANCO=B.BANCO AND A.AGRUPAMIENTO=B.AGRUPAMIENTO AND A.TAXONOMIA=B.TAXONOMIA)", "")
    Call Borrar_Tabla(dbs, "Temp_" & tbl_vf_foto_r)
    
    dbs.Execute "SELECT TAXONOMIA, SUM([MGI (MDP)]) AS MGI_MDP, SUM([SALDO (MDP)]) AS SALDO_MDP, SUM([MPAGADO (MDP)]) AS MPAGADO_MDP, " _
    & " SUM([RECUPERADOS (MDP)]) AS MRECUP_MDP, SUM([RESCATADOS (MDP)]) AS MRESCAT_MDP INTO VALIDA_Foto_R " _
    & " FROM VF_Foto_R_" & Mes & "_VCohortes GROUP BY TAXONOMIA"
       
    dbs.Close
    
    Set dbs = OpenDatabase(db_foto_simples_cohortes)
    
    
    dbs.Execute "SELECT TAXONOMIA, SUM([MGI (MDP)]) AS MGI_MDP, SUM([SALDO (MDP)]) AS SALDO_MDP, SUM([MPAGADO (MDP)]) AS MPAGADO_MDP, " _
    & " SUM([RECUPERADOS (MDP)]) AS MRECUP_MDP, SUM([RESCATADOS (MDP)]) AS MRESCAT_MDP INTO VALIDA_Foto_NR " _
    & " FROM VF_Foto_NR_" & Mes & " GROUP BY TAXONOMIA"

    dbs.Close

End Sub




