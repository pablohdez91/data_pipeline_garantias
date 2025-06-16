Sub Quita_Doble_Destino_Recurso()

    Dim dbs As DAO.database
    Dim dbs_root As DAO.database
    Set Engine = New DBEngine
    'Modificar
    Mes = "202412"
    
    'Declaración de variables
    Mes1 = IIf(Mid(Mes, 5, 1) = 0, Right(Mes, 1), Right(Mes, 2))
    Mes2 = Mes_Letra(Mes1) & Mid(Mes, 3, 2)

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

    db_base = wd_staging & "Base.accdb"
    '''''''''VF_Foto_R_202102_VCohortes & VF_Foto_NR_202102
    
    For i = 1 To 2
        If i = 1 Then
            nr_r = "NR"
            Archivo = wd_processed_fotos_cierre & "FotoSimplesCohortes_" & Mes & "_VF"
            Archivo = wd_processed_fotos_cierre & "FotoSimplesCohortes_" & Mes & "_VF"
            tbl_vf_foto_nr_r = "VF_Foto_NR_" & Mes
        ElseIf i = 2 Then
            nr_r = "R"
            Archivo = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & Mes & "_VF"
            Archivo = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & Mes & "_VF"
            tbl_vf_foto_nr_r = "VF_Foto_R_" & Mes & "_VCohortes"
        End If
        
    ArchivoAux = Archivo & "_Aux"
    tbl_aux_nr_r = "Aux_" & nr_r
    tbl_aux2_nr_r = "Aux2_" & nr_r
    tbl_dobles = "Dobles_" & nr_r
    tbl_valida_foto_vf = "VALIDA_Foto_" & nr_r & "_VF"
    'Se inicia el proceso, se hace todo en el archivo aux
    If ExisteRuta(ArchivoAux) = False Then
       Copia = Application.CompactRepair(db_base, wd_processed_fotos_cierre & "Copia de Seguridad.accdb")
       Name wd_processed_fotos_cierre & "Copia de Seguridad.accdb" As ArchivoAux & ".accdb"
    End If
    Call Vincula_Tabla(Archivo & ".accdb", ArchivoAux & ".accdb", tbl_vf_foto_nr_r, tbl_vf_foto_nr_r)
    Set dbs = OpenDatabase(ArchivoAux)
            dbs.Execute "SELECT *INTO [" & tbl_aux_nr_r & "] FROM [" & tbl_vf_foto_nr_r & "];"
    dbs.Close
    Set dbs = OpenDatabase(ArchivoAux)
        dbs.Execute "SELECT INTER_CLAVE, CLAVE_CREDITO, COUNT(CONREC_CLAVE) AS Dobles " _
        & "INTO [" & tbl_dobles & "] " _
        & "FROM [" & tbl_aux_nr_r & "] " _
        & "GROUP BY INTER_CLAVE, CLAVE_CREDITO;"
    dbs.Close

    Set dbs = OpenDatabase(ArchivoAux)
        dbs.Execute "SELECT A.*, B.Dobles " _
        & "INTO [" & tbl_aux2_nr_r & "] " _
        & "FROM [" & tbl_aux_nr_r & "] A LEFT JOIN [" & tbl_dobles & "] B ON (A.INTER_CLAVE = B.INTER_CLAVE) AND (A.CLAVE_CREDITO = B.CLAVE_CREDITO) "
        dbs.Execute "DROP TABLE [" & tbl_aux_nr_r & "];"
        dbs.Execute "UPDATE [" & tbl_aux2_nr_r & "] SET CONREC_CLAVE = 9999, Describe_Desrec='Doble clave de destino recurso' WHERE  Dobles > 1 "
        
        dbs.Execute "SELECT DISTINCT A.* " _
        & "INTO [" & tbl_aux_nr_r & "] " _
        & "FROM [" & tbl_aux2_nr_r & "] A ; "
        dbs.Execute "ALTER TABLE [" & tbl_aux_nr_r & "] DROP COLUMN Dobles"
    dbs.Close
    
    'Se pega en los archivos de fotos que deben de ser
    Call Vincula_Tabla(ArchivoAux & ".accdb", Archivo & ".accdb", tbl_aux_nr_r, tbl_aux_nr_r)

    
    Set dbs = OpenDatabase(Archivo)
        
        dbs.Execute "DROP TABLE [" & tbl_vf_foto_nr_r & "];"
        dbs.Execute "SELECT *INTO [" & tbl_vf_foto_nr_r & "] FROM [" & tbl_aux_nr_r & "];"

        dbs.Execute "SELECT TAXONOMIA, SUM([MGI (MDP)]) AS MGI, SUM([SALDO (MDP)]) AS SALDO, SUM([MPAGADO (MDP)]) AS MPAGADO, SUM([RECUPERADOS (MDP)]) AS RECUP, SUM([RESCATADOS (MDP)]) AS RESCAT " _
        & "INTO [" & tbl_valida_foto_vf & "] FROM [" & tbl_vf_foto_nr_r & "] GROUP BY TAXONOMIA; "
    dbs.Close
    
    Next
    

End Sub