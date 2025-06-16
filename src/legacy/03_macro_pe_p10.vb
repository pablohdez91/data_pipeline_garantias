'Importar la información que envía Rodrigo Rivera.
Sub Principal_ImportaBase()
    Dim RutaBases As String
    
    MesTextoAct = 202410
    If Right(MesTextoAct, 2) = "01" Then
        MesTextoAnt = (Left(MesTextoAct, 4) - 1) & 12
    Else
        MesTextoAnt = MesTextoAct - 1
    End If


    Mes_P10 = Mes_Letra_Completa(Right(MesTextoAct, 2)) & " " & Left(MesTextoAct, 4)
    Mes_texto = Mes_Letra(Right(MesTextoAct, 2)) & Right(Left(MesTextoAct, 4), 2)
    Mes = Left(Mes_P10, Len(Mes_P10) - 5) & Right(Mes_P10, 2)
    
    wd = "D:\DAR\proyecto_mejora_fotos\2. Nuevas fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_dwh = wd_processed & "DWH\"
    wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
    wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"

    'Input
    db_catalogos = wd_external & "Catálogos_" & mes & ".accdb"
    db_base_vigente_xl = wd_external & "Base Vigente " & Mes_P10 & ".xlsx"

    'Output
    db_base_vigente = wd_staging & "Base Vigente " & Mes & ".accdb"

    tbl_temp = Mes & "_Temp"

    'Crea Archivo Base
    Set Engine = New DBEngine
    Set dbs_root = Engine.CreateDatabase(db_base_vigente, dbLangGeneral)
    dbs_root.Close

    Call importar_hoja_excel(db_base_vigente_xl, db_base_vigente, tbl_temp)
    Call importar_catalogos(db_base_vigente, Mes, db_catalogos)
    Call pegar_tc(db_base_vigente, Mes)
    Call pegar_catalogos(db_base_vigente, Mes)
    Call pegar_catalogos_valida_saldo(db_base_vigente, Mes)
    Set dbs = OpenDatabase(db_base_vigente)
        Call corrige_campos(dbs, "Saldos_" & Mes, "Taxonomia", "IIF((Agrupamiento_Id=655 AND Agrupamiento = 'Mis' and [Taxonomia]='GARANTIA MICROCREDITO'),'GARANTIA SECTORIAL',[Taxonomia])", "")
        Call corrige_campos(dbs, "valida_saldos_" & Mes, "Taxonomia", "IIF((Agrupamiento_Id=655 AND Agrupamiento = 'Mis' and [Taxonomia]='GARANTIA MICROCREDITO'),'GARANTIA SECTORIAL',[Taxonomia])", "")
        Call valida_saldo(dbs, Mes)
        Call valida_agrupamientos(dbs, Mes)
    dbs.Close
End Sub




Function corrige_campos(dbs, TablaMod, Columna, Condicion, Filtro)
    dbs.Execute "update [" & TablaMod & "] " _
    & "set [" & Columna & "] = " & Condicion & " " _
    & "" & Filtro & " ;"
End Function

Function importar_catalogos(ArchivoDestino, Mes, db_catalogos)
    Set dbs = OpenDatabase(ArchivoDestino)
        dbs.Execute "SELECT * INTO BANCO_" & Mes & " FROM BANCO IN '" & db_catalogos & "';"
        dbs.Execute "SELECT * INTO AGRUPAMIENTO_" & Mes & " FROM AGRUPAMIENTO IN '" & db_catalogos & "';"
        dbs.Execute "SELECT * INTO PROGRAMA_" & Mes & " FROM PROGRAMA IN '" & db_catalogos & "';"
        dbs.Execute "SELECT * INTO GARANTIA_" & Mes & " FROM GARANTIA IN '" & db_catalogos & "';"
        dbs.Execute "SELECT * INTO TIPO_CREDITO_" & Mes & " FROM TIPO_CREDITO IN '" & db_catalogos & "';"
    dbs.Close
End Function

Function importar_hoja_excel(BaseOrigen, BaseDestino, TablaDestino)
    'BaseOrigen = "G:\INFO_NAFIN\GARANTIAS\COHORTES\DWH\2011\201105\Recuperadas_con_Pagos_Flujos_y_Gastos_Juicio_May11.xlsx"
    Dim objetAccessO As Access.Application
    Set objetAccessO = New Access.Application
    objetAccessO.OpenCurrentDatabase BaseDestino
    objetAccessO.DoCmd.TransferSpreadsheet acImport, 9, TablaDestino, BaseOrigen, True
    objetAccessO.Quit
    Set objetAccessO = Nothing
End Function

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

Function Mes_Letra_Completa(Mes) As String
    Select Case Mes
    Case 1
        Mes_Letra_Completa = "Enero"
    Case 2
        Mes_Letra_Completa = "Febrero"
    Case 3
        Mes_Letra_Completa = "Marzo"
    Case 4
        Mes_Letra_Completa = "Abril"
    Case 5
        Mes_Letra_Completa = "Mayo"
    Case 6
        Mes_Letra_Completa = "Junio"
    Case 7
        Mes_Letra_Completa = "Julio"
    Case 8
        Mes_Letra_Completa = "Agosto"
    Case 9
        Mes_Letra_Completa = "Septiembre"
    Case 10
        Mes_Letra_Completa = "Octubre"
    Case 11
        Mes_Letra_Completa = "Noviembre"
    Case 12
        Mes_Letra_Completa = "Diciembre"
    End Select
 End Function

Function pegar_catalogos(ArchivoDestino, Mes)
Dim dbs As DAO.Database
Set dbs = OpenDatabase(ArchivoDestino)
dbs.Execute "SELECT I.*, D.AGRUPAMIENTO " _
          & "INTO [Saldos_" & Mes & "] " _
          & "FROM (SELECT J.*, F.Producto AS Taxonomia " _
          & "FROM (SELECT C.*, IIF(H.ESQUEMA='1P','1P','PP') AS ESQUEMA, H.AGRUPAMIENTO_ID " _
          & "FROM (SELECT A.*, IIF(LEN(A.[tpro_clave])<=3,(CLng(3000)+A.[tpro_clave]),(clng(30000)+A.[tpro_clave])) as Programa_Id, " _
          & "clng(NZ(IIF(LEN(A.[tpro_clave_original])<=3,(clng(3000)+ A.[tpro_clave_original]),(clng(30000)+A.[tpro_clave_original])),clng(3999))) as Progama_Original, " _
          & "B.Intermediario_Id, B.[Razón Social] " _
          & "FROM [" & Mes & "] A LEFT JOIN BANCO_" & Mes & " B " _
          & " ON A.INTER_CLAVE=B.Intermediario_ID_SIAG " _
          & " WHERE A.[descripcion portafolio] NOT IN ('SUBASTA','SELECTIVA','SELECTIVAS')) C LEFT JOIN PROGRAMA_" & Mes & " H " _
          & " ON C.Programa_Id=H.PROGRAMA_ID) J LEFT JOIN GARANTIA_" & Mes & " F " _
          & " ON J.[descripcion portafolio]=F.Programa_SIAG) I LEFT JOIN AGRUPAMIENTO_" & Mes & " D " _
          & " ON I.AGRUPAMIENTO_ID=D.AGRUPAMIENTO_ID; "
dbs.Close
End Function

Function pegar_catalogos_valida_saldo(ArchivoDestino, Mes)
Dim dbs As DAO.Database
Set dbs = OpenDatabase(ArchivoDestino)
dbs.Execute "SELECT I.*, D.Agrupamiento " _
          & "INTO [valida_saldos_" & Mes & "] " _
          & "FROM (SELECT H.*, F.Producto AS Taxonomia " _
          & "FROM (SELECT C.*, IIF(H.ESQUEMA='1P','1P','PP') AS ESQUEMA, H.Agrupamiento_Id " _
          & "FROM (SELECT A.*, IIF(LEN(A.[tpro_clave])<=3,(CLng(3000)+A.[tpro_clave]),(clng(30000)+A.[tpro_clave])) as Programa_Id, " _
          & "clng(NZ(IIF(LEN(A.[tpro_clave_original])<=3,(clng(3000)+ A.[tpro_clave_original]),(clng(30000)+A.[tpro_clave_original])),clng(3999))) as Progama_Original, " _
          & "B.Intermediario_Id, B.[Razón Social] " _
          & "FROM [" & Mes & "] A LEFT JOIN BANCO_" & Mes & " B " _
          & " ON A.INTER_CLAVE=B.Intermediario_ID_SIAG) C LEFT JOIN PROGRAMA_" & Mes & " H  " _
          & "ON C.Programa_Id=H.PROGRAMA_ID) H LEFT JOIN GARANTIA_" & Mes & " F " _
          & "ON H.[descripcion portafolio]=F.Programa_SIAG) I LEFT JOIN AGRUPAMIENTO_" & Mes & " D " _
          & "ON I.AGRUPAMIENTO_ID=D.AGRUPAMIENTO_ID;"
dbs.Close
End Function

Function pegar_tc(ArchivoDestino, Mes)
Dim dbs As DAO.Database
Set dbs = OpenDatabase(ArchivoDestino)
dbs.Execute "SELECT A.*, B.Tipo_Credito_Id " _
          & "INTO [" & Mes & "] " _
          & "FROM [" & Mes & "_Temp] as A LEFT JOIN [TIPO_CREDITO_" & Mes & "] as B " _
          & "ON A.[ticr_clave]=B.[Clave (Garantias)]; "
dbs.Close
End Function

 Function valida_agrupamientos(dbs, Mes)
'Cambié Saldo_contingente_mn a SALDO GARANTIZADO por cambio en el archivo de Chris 201508 ACOV
dbs.Execute "SELECT A.Programa_Id, A.Progama_Original, A.Taxonomia, Sum(A.[SALDO GARANTIZADO]) AS SumaDeSaldo_contingente_mn, A.Agrupamiento " _
          & "INTO [BD_Sin_Agrupamiento_" & Mes & "] " _
          & "FROM [valida_saldos_" & Mes & "] A  " _
          & "GROUP BY A.Programa_Id, A.Progama_Original, A.Taxonomia, A.Agrupamiento " _
          & "HAVING (((A.Agrupamiento) Is Null));"
End Function

Function valida_saldo(dbs, Mes)
'Cambié Saldo_contingente_mn a SALDO GARANTIZADO por cambio en el archivo de Chris 201508 ACOV
dbs.Execute "SELECT A.[descripcion portafolio], A.inter_clave, A.Intermediario_Id, A.[Razón Social], A.tpro_clave_original, A.tpro_clave, A.Progama_Original, A.Programa_Id, Sum(A.[SALDO GARANTIZADO]) AS SumaDeSaldo_contingente_mn " _
          & "INTO [Consulta_valida_saldos_" & Mes & "] " _
          & "FROM [valida_saldos_" & Mes & "] A  " _
          & " GROUP BY A.[descripcion portafolio], A.inter_clave, A.Intermediario_Id, A.[Razón Social], A.tpro_clave_original, A.tpro_clave, A.Progama_Original, A.Programa_Id; "
End Function