Sub Principal_Foto()
    'Datos a Modificar
    Dim dbs As DAO.database
    Dim dbs_root As DAO.database
    Dim Ruta As String
    Set Engine = New DBEngine
    'Modificar
    Var_NR_R = "R"
    Mes = "202412"
    
    Mes1 = IIf(Mid(Mes, 5, 1) = 0, Right(Mes, 1), Right(Mes, 2))
    Mes2 = Mes_Letra(Mes1) & Mid(Mes, 3, 2)

    wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_dwh = wd_processed & "DWH\"
    wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
    wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
    wd_fotos = wd_processed & "Fotos\"
    wd_processed_curvarecup = wd_processed & "CurvaRecup\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"
    
    db_catalogos = wd_external & "Catálogos_" & Mes & ".accdb"
    db_base = wd_staging & "Base.accdb"
    db_base_vacia = wd_staging & "Base_Vacia.accdb"

    db_recupera_con_pagos_flujos_finales = wd_processed_dwh_bases_finales & "Recupera_con_Pagos_Flujos_" & Mes & ".accdb"
    db_querie_curva_recuperada = wd_processed_curvarecup & "Querie_CurvaRecuperada_" & Mes & ".accdb"
    db_acumulado_saldos = wd_fotos & "Acumulado_Saldos_" & Left(Mes, 4) & ".accdb"

    db_bd_dwh_vf = wd_fotos & "BD_DWH_" & Mes & "_VF.accdb"

    tbl_pagadas_detalle_vf = "PAGADAS_DETALLE_VF_" & Mes  '"Pagadas_Ult_Detalle_" & Mes2
    tbl_recupera_con_pagos_flujos = "Recupera_con_Pagos_Flujos_" & Mes

    Foto = "BD_DWH_" & Var_NR_R & "_" & Mes
    TRO = "Recuperadas_Global_VF_" & Mes2

    tbl_bd_dwh = "BD_DWH_" & Var_NR_R & "_" & Mes
    tbl_recuperadas_global_vf = "Recuperadas_Global_VF_" & Mes2

    'Define Ruta Simpls o Revolventes
    If Var_NR_R = "NR" Then
        db_simples_revolventes = wd_fotos & "\Simples_" & Mes & ".accdb"
        Carpeta = "Simples"
    ElseIf Var_NR_R = "R" Then
        db_simples_revolventes = wd_fotos & "\Revolventes_" & Mes & ".accdb"
        Carpeta = "Revolventes"
    End If
    'BaseDestinoFoto = "G:\INFO_NAFIN\GARANTIAS\COHORTES\" & Carpeta & "\" & Left(Mes, 4) & "\" & Mes & "\Maquetas " & var_nr_r & "\1 Importa Foto " & var_nr_r & ".accdb"
    If ExisteRuta(wd_fotos) = False Then
        MkDir (wd_fotos)
    End If
    
    If ExisteRuta(db_simples_revolventes) = False Then
        Copia = Application.CompactRepair(db_base, wd_fotos & "\Copia de Seguridad.accdb")
        Name wd_fotos & "\Copia de Seguridad.accdb" As db_simples_revolventes
    End If
    
    'Importar_TXT_a_BDAccess db_simples_revolventes, BD_KATALOGO_UDI, "UDI", "KATALOGO_UDI"
    'Datos RecupCohortes
     'BDH_Catalogo
      
     Linea = "A.Numero_Credito, A.Intermediario_Id, A.NR_R, A.Producto"
     tbl_z3_recup_cohor = "Z3_RECUPCOHOR"
     'Foto = "BD_DWH_" & Mes '"Foto_" & var_nr_r & "_" & Mes & "_DWH"
     tbl_vf_foto_nr_r = "VF_Foto_" & Var_NR_R & "_" & Mes
     tbl_vf_recuperaciones = "VF_Recuperaciones_" & Mes
     filtro_smicro = "and ( A.[TAXONOMIA]<>'GARANTIA MICROCREDITO' and A.[TAXONOMIA] <> 'GARANTIAS BURSATIL' and A.[TAXONOMIA] <> 'GARANTIAS BANCOMEXT' and A.[TAXONOMIA] <>'GARANTIAS PRIMER PISO'))"
     filtro_micro = "and ( A.[TAXONOMIA]='GARANTIA MICROCREDITO' and A.[TAXONOMIA] <> 'GARANTIAS BURSATIL' and A.[TAXONOMIA] <> 'GARANTIAS BANCOMEXT' and A.[TAXONOMIA] <>'GARANTIAS PRIMER PISO'))"

    'Datos Genera_Saldos
     tbl_saldos = "Saldos " & Mes
    'Vinculo Foto a Base Origen
     Vincula_Tabla db_bd_dwh_vf, db_simples_revolventes, tbl_bd_dwh, tbl_bd_dwh
     Vincula_Tabla db_recupera_con_pagos_flujos_finales, db_simples_revolventes, tbl_recupera_con_pagos_flujos, tbl_recuperadas_global_vf
    
'**************aquiiiiiiiiii *******************Se agregan Pagadas de la tabla de Pagadas************************************************************************************************************************************************************
    
    'Vincula Pagadas MMMAA
    Vincula_Tabla db_querie_curva_recuperada, db_simples_revolventes, tbl_pagadas_detalle_vf, tbl_pagadas_detalle_vf
    'Paso2 db_simples_revolventes, Linea, tbl_z3_recup_cohor, tbl_vf_recuperaciones
    Une_Pagadas Var_NR_R, Mes, db_simples_revolventes, "Temp_" & tbl_bd_dwh, tbl_bd_dwh, tbl_pagadas_detalle_vf

    'Elimina Tabla anterior
    Borrar_Tabla_BO db_simples_revolventes, tbl_bd_dwh
    'Renombra Tabla
    RenombraTablas db_simples_revolventes, "Temp_" & tbl_bd_dwh, tbl_bd_dwh
    
    'Llama a módulo RecupCohortes
     Z3_RECUPCOHORTES_NR Var_NR_R, Mes, db_simples_revolventes, Linea, tbl_recuperadas_global_vf, tbl_z3_recup_cohor, tbl_bd_dwh, tbl_vf_recuperaciones

     Paso2 db_simples_revolventes, Linea, tbl_z3_recup_cohor, tbl_vf_recuperaciones
    
    'Elimina Tabla anterior
     Borrar_Tabla_BO db_simples_revolventes, tbl_z3_recup_cohor

    'Tabla_Katalog db_simples_revolventes, Mes
    Foto_Saldo db_simples_revolventes, tbl_bd_dwh, "VF_Pagadas_" & Var_NR_R & "_" & Mes, Mes, Var_NR_R ', Filtro_Completo, Katalogo_UDI, Katalogo_Programa, Katalogo_Agrupamiento   'Para Revolventes y Simple completa

'*********************************Se agregan Recuperaciones de la tabla de Recuperaciones**************
    'Se Crea un Libro por separado para los resumenes Finales Debido al Tamaño
    If Var_NR_R = "NR" Then
        db_foto_simples_revolventes_cohortes = wd_fotos & "\FotoSimplesCohortes_" & Mes & "_VF.accdb"
        
    Else
        db_foto_simples_revolventes_cohortes = wd_fotos & "\FotoRevolventesCohortes_" & Mes & "_Preeliminar.accdb"
    '    tbl_vf_pagadas_nr_r = "VF_Pagadas_" & var_nr_r & "_" & Mes
    End If
    tbl_vf_pagadas_nr_r = "VF_Pagadas_" & Var_NR_R & "_" & Mes
    Set dbs = Engine.CreateDatabase(db_foto_simples_revolventes_cohortes, dbLangGeneral)
    dbs.Close
    
    'Genera Foto Completa
    Une_Pagos_Vs_Recup Var_NR_R, Mes, db_simples_revolventes, tbl_vf_foto_nr_r, tbl_vf_pagadas_nr_r, tbl_vf_recuperaciones, "", db_foto_simples_revolventes_cohortes     'Para Revolventes y Simple completa

     'Genera 2 Fotos Para Simples  'Simples_Micro' y 'Simples_sin_Micro'
     If Var_NR_R = "NR" Then
        filtro_resto = " Where ( A.[TAXONOMIA]<>'GARANTIA MICROCREDITO' AND A.[TAXONOMIA]<>'GARANTIA EMPRESARIAL' AND A.[TAXONOMIA]<>'GARANTIAS BURSATIL' AND A.[TAXONOMIA]<>'GARANTIAS BANCOMEXT' AND A.[TAXONOMIA]<>'GARANTIAS PRIMER PISO' AND A.[TAXONOMIA]<>'GARANTIAS SHF/LI FINANCIERO') "  '"Where ( A.[TAXONOMIA] Not In ('GARANTIA MICROCREDITO','GARANTIAS BURSATIL','GARANTIAS BANCOMEXT','GARANTIAS PRIMER PISO')"
        filtro_empresarial = " Where ( A.[TAXONOMIA]='GARANTIA EMPRESARIAL') "  '"Where ( A.[TAXONOMIA] Not In ('GARANTIA MICROCREDITO','GARANTIAS BURSATIL','GARANTIAS BANCOMEXT','GARANTIAS PRIMER PISO')"
        filtro_micro = " Where ( A.[TAXONOMIA]='GARANTIA MICROCREDITO') "    '"Where ( A.[TAXONOMIA]='GARANTIA MICROCREDITO' and A.[TAXONOMIA] NOT IN ('GARANTIAS BURSATIL','GARANTIAS BANCOMEXT','GARANTIAS PRIMER PISO')"
        Une_Pagos_Vs_Recup Var_NR_R, Mes, db_simples_revolventes, tbl_vf_foto_nr_r & "_Resto", "VF_Pagadas_" & Var_NR_R & "_" & Mes, tbl_vf_recuperaciones, filtro_resto, db_foto_simples_revolventes_cohortes          'Simple Empresarial
        Une_Pagos_Vs_Recup Var_NR_R, Mes, db_simples_revolventes, tbl_vf_foto_nr_r & "_Empresarial", "VF_Pagadas_" & Var_NR_R & "_" & Mes, tbl_vf_recuperaciones, filtro_empresarial, db_foto_simples_revolventes_cohortes          'Simple Resto
        Une_Pagos_Vs_Recup Var_NR_R, Mes, db_simples_revolventes, tbl_vf_foto_nr_r & "_Micro", "VF_Pagadas_" & Var_NR_R & "_" & Mes, tbl_vf_recuperaciones, filtro_micro, db_foto_simples_revolventes_cohortes        'Simple Micro
     End If
     'Llama a módulo Genera_Saldos
     Genera_BaseSaldos db_foto_simples_revolventes_cohortes, Mes, tbl_vf_foto_nr_r, tbl_saldos
     Call Exporta_Saldos(wd_fotos, db_acumulado_saldos, Mes, tbl_saldos)
     
End Sub


Sub Exporta_Saldos(wd_fotos, db_acumulado_saldos, Mes, tbl_saldos)
    'Pregunta si existen archivos, si si hace la exportacion, si no ent manda msg de que aun no lo ha hecho xq falta algun archivo sale
    If ExisteRuta(wd_fotos & "\Simples_" & Mes & ".accdb") = False Or ExisteRuta(wd_fotos & "\Revolventes_" & Mes & ".accdb") = False Then
        MsgBox "Aun no se ha creado la Tabla de Saldos: " & tbl_saldos
        Exit Sub
    End If
    If ExisteTabla(db_acumulado_saldos, tbl_saldos) = True Then
        Set dbs = OpenDatabase(db_acumulado_saldos)
        Borrar_Tabla dbs, tbl_saldos
        dbs.Close
    End If
    CopiarTabla_BD wd_fotos & "\FotoSimplesCohortes_" & Mes & "_VF.accdb", db_acumulado_saldos, tbl_saldos, tbl_saldos
    Set dbs = OpenDatabase(db_acumulado_saldos)
        Inserta_Filas dbs, " IN '" & wd_fotos & "\FotoRevolventesCohortes_" & Mes & "_Preeliminar.accdb" & "'", "*", tbl_saldos, tbl_saldos, "", ""
    dbs.Close
End Sub


 Function Mes_Letra(Mes) As String
    Select Case Mes
    Case 1
        Mes_Letra = "Enero"
    Case 2
        Mes_Letra = "Febrero"
    Case 3
        Mes_Letra = "Marzo"
    Case 4
        Mes_Letra = "Abril"
    Case 5
        Mes_Letra = "Mayo"
    Case 6
        Mes_Letra = "Junio"
    Case 7
        Mes_Letra = "Julio"
    Case 8
        Mes_Letra = "Agosto"
    Case 9
        Mes_Letra = "Septiembre"
    Case 10
        Mes_Letra = "Octubre"
    Case 11
        Mes_Letra = "Noviembre"
    Case 12
        Mes_Letra = "Diciembre"
    End Select
    Mes_Letra = Left(Mes_Letra, 3)
 End Function
 

