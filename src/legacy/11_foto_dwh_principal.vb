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
 
Sub Une_Pagadas(Var_NR_R, Mes, BaseOrigen, FotoFinal, Pagos_VF, T_Pagadas)
    Dim dbs As DAO.database
    Set dbs = OpenDatabase(BaseOrigen)
    dbs.Execute "SELECT A.*, (IIF(B.[Monto_Desembolso_Mn] is NULL, 0,B.[Monto_Desembolso_Mn])+IIF(B.[Interes_Desembolso_Mn] is NULL, 0,B.[Interes_Desembolso_Mn])+IIF(B.[Interes_Moratorios_Mn] is NULL, 0,B.[Interes_Moratorios_Mn]))/1000000 as [MPAGADO (MDP)], " _
        & "IIF((IIF(B.[Monto_Desembolso_Mn] is NULL, 0,B.[Monto_Desembolso_Mn])+IIF(B.[Interes_Desembolso_Mn] is NULL, 0,B.[Interes_Desembolso_Mn])+IIF(B.[Interes_Moratorios_Mn] is NULL, 0,B.[Interes_Moratorios_Mn]))>0,1,0) as PAGADAS, " _
        & "IIF((IIF(B.[Monto_Desembolso_Mn] is NULL, 0,B.[Monto_Desembolso_Mn])+IIF(B.[Interes_Desembolso_Mn] is NULL, 0,B.[Interes_Desembolso_Mn])+IIF(B.[Interes_Moratorios_Mn] is NULL, 0,B.[Interes_Moratorios_Mn]))>0,1,0) as INCUMPLIDO, " _
        & "IIF(B.[Fecha_Garantia_Honrada] is NULL,cdate(Format('30/12/1899','dd/mm/yyyy')), B.[Fecha_Garantia_Honrada]) as FECHA_PAGO " _
        & "INTO [" & FotoFinal & "] " _
        & "FROM [" & Pagos_VF & "] A LEFT JOIN [" & T_Pagadas & "] B " _
        & "on (cstr(A.Intermediario_Id)=cstr(B.Intermediario_Id) AND A.Numero_Credito=B.Numero_Credito) " _
        & "IN '" & BaseOrigen & "'; "
    dbs.Close
End Sub

Function Borrar_Tabla_BO(BaseOrigen, Tabla)
    Set dbs = OpenDatabase(BaseOrigen)
    dbs.Execute "DROP TABLE [" & Tabla & "] ;"
    dbs.Close
End Function

Function RenombraTablas(RutaBaseDatosOrigen, TablaOrigen, TablaNueva)
     Dim objetAccessO As Access.Application
     Set objetAccessO = New Access.Application
     objetAccessO.OpenCurrentDatabase RutaBaseDatosOrigen
     objetAccessO.DoCmd.Rename TablaNueva, acTable, TablaOrigen
     objetAccessO.Quit
     Set objetAccessO = Nothing
End Function

Function Z3_RECUPCOHORTES_NR(Var_NR_R, Mes, BaseOrigen, Linea, TRO, TablaZ3, Foto, RecuperacionesVF)
    Dim dbs As database
    ' Z3_RECUPCOHORTES_NR Var_NR_R, Mes, db_simples_revolventes, Linea, tbl_recuperadas_global_vf, tbl_z3_recup_cohor, tbl_bd_dwh, tbl_vf_recuperaciones 
    TRO_Temp = TRO & "_Origen_Temp"
    Set dbs = OpenDatabase(BaseOrigen)
    dbs.Execute "SELECT N.*, N.Numero_Credito & N.Intermediario_Id as Concatenado " _
        & "into [" & TRO_Temp & "] " _
        & "FROM [" & TRO & "] as N " _
        & "IN '" & BaseOrigen & "'; "
   
    dbs.Execute "SELECT " & Linea & ", " _
        & "IIF (A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]), 1, 0) AS ENTRA_RECUP, " _
        & "IIF ((A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]) AND (A.Estatus='D' or A.Estatus='E' or A.Estatus='RI' or A.Estatus='CR' or A.Estatus='RAR' or A.Estatus='RAC' or A.Estatus='CJ' or A.Estatus='CS' or A.Estatus='R' or A.Estatus='RJ' or A.Estatus='RS')), (nz(A.Monto_Mn ,0)+nz(A.Interes_Mn,0)+nz(A.Moratorios_Mn,0)+nz(A.Excedente_Mn,0)-nz(A.[Gastos_Juicio_Mn],0))/1000000, 0) AS [MONTOTOTAL (MDP)]," _
        & "IIF ((A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]) AND (A.Estatus='D' or A.Estatus='E' or A.Estatus='RI' or A.Estatus='CR' or A.Estatus='RAR' or A.Estatus='RAC')), (nz(A.Monto_Mn,0)+nz(A.Interes_Mn,0)+nz(A.Moratorios_Mn,0)+nz(A.Excedente_Mn,0)-nz(A.[Gastos_Juicio_Mn],0))/1000000,0) AS [RECUPERADOS (MDP)], " _
        & "IIF ((A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]) AND (A.Estatus='CJ' or A.Estatus='CS' or A.Estatus='R' or A.Estatus='RJ' or A.Estatus='RS')), (nz(A.Monto_Mn,0)+nz(A.Interes_Mn,0)+nz(A.Moratorios_Mn,0)+nz(A.Excedente_Mn,0)-nz(A.[Gastos_Juicio_Mn],0))/1000000,0) AS [RESCATADOS (MDP)] " _
        & "INTO [" & TablaZ3 & "] " _
        & "FROM [" & TRO_Temp & "] as A LEFT JOIN (SELECT M.*,  M.[Numero_Credito] & M.Intermediario_Id as Concatenado2 FROM [" & Foto & "] M) as B " _
        & "on (A.Concatenado=B.Concatenado2) " _
        & "IN '" & BaseOrigen & "'; "
    Borrar_Tabla dbs, TRO_Temp
    dbs.Close
End Function

Function Paso2(BaseOrigen, Linea, TablaZ3, RecuperacionesVF)
Dim dbs As database
Set dbs = OpenDatabase(BaseOrigen)
dbs.Execute "SELECT " & Linea & ", " _
    & "SUM([A.MONTOTOTAL (MDP)]) AS [MONTOTOTAL (MDP)], " _
    & "SUM([A.RECUPERADOS (MDP)]) AS [RECUPERADOS (MDP)], " _
    & "SUM([A.RESCATADOS (MDP)]) AS [RESCATADOS (MDP)] " _
    & "into [" & RecuperacionesVF & "] " _
    & "FROM [" & TablaZ3 & "] as A " _
    & "IN '" & BaseOrigen & "' " _
    & "GROUP BY " & Linea & " " _
    & "ORDER BY " & Linea & "; "
dbs.Close
End Function

Function Foto_Saldo(BaseOrigen, Saldos_mes_Inicial, Saldos_mes_Final, Mes, Var_NR_R)  ', Filtro, Katalogo_UDI, Katalogo_Programa, Katalogo_Agrupamiento
    Dim dbs As DAO.database
    Set dbs = OpenDatabase(BaseOrigen)
    dbs.Execute "SELECT  A.BUCKET, A.CAMBIO, A.[Monto _Credito_Mn]*A.CAMBIO AS MCrédito_MM_UDIS, A.[MM_UDIS], " _
        & "A.[Intermediario_Id] as INTER_CLAVE, A.Nombre_v1 as NOMBRE, A.[RFC Empresa / Acreditado] as RFC, A.[TIPO_PERSONA] as TIPO_PERSONA, A.[Numero_Credito] as CLAVE_CREDITO, " _
        & "A.[Fecha de Apertura] as FECHA_VALOR, IIF(A.[Plazo Días] IS NULL,0,A.[Plazo Días]) as PLAZO_DIAS, A.[Plazo] as PLAZO, A.[FVTO_Riesgosd] as FVTO, A.[Fecha Registro Alta] as FECHA_REGISTRO_GARANTIA, " _
        & "A.[Monto_Garantizado_Mn]/1000000 as [MGI (MDP)], A.[Porcentaje Garantizado] as PORCENTAJE_GARANTIZADO, A.[Razón Social (Intermediario)] as BANCO, IIF(A.[Fecha_Primer_Incumplimiento] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),A.[Fecha_Primer_Incumplimiento]) as FECHA_PRIMER_INCUM, " _
        & "A.[Monto _Credito_Mn]/1000000 as [MONTO CREDITO (MDP)], A.[Saldo_Contingente_Mn]/1000000 as [SALDO (MDP)], A.[TPRO_CLAVE] as TPRO_CLAVE, " _
        & "A.[Producto ID] as CLAVE_TAXO, A.[Producto] as TAXONOMIA, A.[NR_R], " _
        & "IIF(A.[Fecha de Apertura]=0,NULL, cdate(Format(dateserial(Year(A.[Fecha de Apertura]),Month(A.[Fecha de Apertura]),'01'),'dd/mm/yyyy'))) AS FECHA_VALOR1, " _
        & "IIF(A.[Fecha Registro Alta]=0,NULL, cdate(Format(dateserial(Year(A.[Fecha Registro Alta]),Month(A.[Fecha Registro Alta]),'01'),'dd/mm/yyyy'))) AS FECHA_REGISTRO1, " _
        & "IIF(A.[Numero_Credito] is NULL, 0,1) AS NUM_GAR, A.[CSG], " _
        & "IIF(Plazo<=12,1,IIF(Plazo<=24,2,IIF(Plazo<=36,3,4))) AS PLAZO_BUCKET, A.[MPAGADO (MDP)], A.PAGADAS, A.INCUMPLIDO, A.FECHA_PAGO, " _
        & "A.[Programa_Original] as Programa_Original, A.[Programa_Id] as Programa_Id, A.[Estrato_Id] as Estrato_Id, A.[Sector_Id] as Sector_Id, A.[Estado_Id] as Estado_Id, A.[Tipo_Credito_Id] as Tipo_Credito_Id, A.[Porcentaje de Comisión Garantia] as Porcentaje_Comision_Garantia, " _
        & "A.[Tasa_Id] as Tasa_Id, A.[Valor_Tasa_Interes] as [Tasa_Interes],  A.[Monto_Garantizado_Mn_Original]/1000000 as [MGI (MDP) Original], A.[AGRUPAMIENTO_ID], " _
        & "A.ESQUEMA, A.SUBESQUEMA, A.AGRUPAMIENTO, A.FONDOS_CONTRAGARANTIA, A.CONREC_CLAVE, A.Describe_Desrec " _
        & "INTO " & Saldos_mes_Final & "  " _
        & "FROM [" & Saldos_mes_Inicial & "] A " _
        & " ;"
End Function

Sub Une_Pagos_Vs_Recup(Var_NR_R, Mes, BaseOrigen, FotoFinal, Pagos_VF, Recuperaciones, Filtro, BaseDestino)
    Dim dbs As DAO.database
    Set dbs = OpenDatabase(BaseDestino)
    dbs.Execute "SELECT A.*, B.[MONTOTOTAL (MDP)], B.[RECUPERADOS (MDP)], B.[RESCATADOS (MDP)] " _
              & "INTO [" & FotoFinal & "] " _
              & "FROM " & Pagos_VF & " AS A LEFT JOIN " & Recuperaciones & " AS B " _
              & " ON (cstr(A.INTER_CLAVE)=cstr(B.Intermediario_Id) AND A.CLAVE_CREDITO=B.Numero_Credito) " _
              & " IN '" & BaseOrigen & "' " _
              & " " & Filtro & " ;"
    dbs.Close
End Sub

Function Genera_BaseSaldos(BaseOrigen, Mes, FotoFinal, SaldoFinal)
Dim dbs As DAO.database
Set dbs = OpenDatabase(BaseOrigen)
dbs.Execute "SELECT A.BUCKET, A.INTER_CLAVE, A.CLAVE_CREDITO, A.BANCO, A.[SALDO (MDP)] AS SALDO_MDP, A.CLAVE_CREDITO & A.INTER_CLAVE AS CONCATENAR_SALDOS " _
          & "INTO [" & SaldoFinal & "] " _
          & "FROM [" & FotoFinal & "] AS A " _
          & "WHERE (A.[SALDO (MDP)]) > 0; "
dbs.Close
End Function