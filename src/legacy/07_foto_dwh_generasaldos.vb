Sub GeneraSaldosActual()
    Dim Engine As DBEngine
    Dim dbs As DAO.database
    Set Engine = New DBEngine
    Mes = "202412"

    wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_fotos = wd_processed & "Fotos\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"

    db_dwh_inter = wd_processed_fotos & "BD_DWH_" & Mes & "_Inter.accdb"
    db_dwh_vf = wd_processed_fotos & "BD_DWH_" & Mes & "_VF.accdb"

    tbl_bd_dwh_saldoymgi_vf = "BD_DWH_SaldoyMGI_" & Mes & "_VF"

    If ExisteRuta(db_dwh_vf) = False Then
        Set dbs = Engine.CreateDatabase(db_dwh_vf, dbLangGeneral)
        dbs.Close
    End If


    For k = 1 To 2
        If k = 1 Then
           Var_NR_R = "R"
        Else
           Var_NR_R = "NR"
        End If
        db_dwh_nr_r = "BD_DWH_" & Var_NR_R & "_" & Mes
        db_dwh_nr_r_completa = "BD_DWH_" & Var_NR_R & "_" & Mes & "_Completa"
        Set dbs = OpenDatabase(db_dwh_vf)
        dbs.Execute "SELECT A.*, IIF(B.[Moneda_id] = 54, A.[TC] * B.[Saldo_Contingente],B.[Saldo_Contingente]) as Saldo_Contingente_Mn, " _
            & "IIF(B.[Moneda_id] = 54, A.[TC] * B.[Monto_Garantizado],B.[Monto_Garantizado]) as Monto_Garantizado_Mn_Original, B.Moneda_id, " _
            & "(A.[Monto _Credito_Mn]*A.[Porcentaje Garantizado]/100) as Monto_Garantizado_Mn " _
            & "INTO [" & db_dwh_nr_r & "] " _
            & "FROM [" & db_dwh_nr_r_completa & "] A LEFT JOIN [" & tbl_bd_dwh_saldoymgi_vf & "] B " _
            & "ON (A.Intermediario_Id & A.Numero_Credito=B.Concatenado) " _
            & "in '" & db_dwh_inter & "';"
            Call Corrige_Campos(dbs, db_dwh_nr_r, "TC", 1, "Where [Moneda_id] = 1")
        dbs.Close
        'Call Vincula_Tabla(BDH_Catalogo, db_dwh_inter, "TIPO_GARANTIA", "Katalogo_Garantia")
        Call Compara_Registros(db_dwh_vf, db_dwh_inter, db_dwh_nr_r, db_dwh_nr_r_completa, "Nombre_v1", "Nombre_v1")
    
    Next
    
    'Genera validación agrupamiento y CSF ACOV201803
    
    'Agrupamiento
    Set dbs = OpenDatabase(db_dwh_vf)
    dbs.Execute "Select Producto, SUM(Saldo_Contingente_Mn) as Saldo_MDP_Suma INTO Valida_Agrup_NR FROM BD_DWH_NR_" & Mes & " WHERE AGRUPAMIENTO IS NULL GROUP BY Producto"
    dbs.Execute "Select Producto, SUM(Saldo_Contingente_Mn) as Saldo_MDP_Suma INTO Valida_Agrup_R FROM BD_DWH_R_" & Mes & " WHERE AGRUPAMIENTO IS NULL GROUP BY Producto"
    dbs.Close
    
    'CSF solo la vincula
    Set dbs = OpenDatabase(db_dwh_vf)
    Mes1 = IIf(Mid(Mes, 5, 1) = 0, Right(Mes, 1), Right(Mes, 2))
    Mes2 = Mes_Letra(Mes1) & Mid(Mes, 3, 2)

    db_catalogos = wd_external & "Catálogos_" & Mes2 & ".accdb"

    tbl_sin_fondos_contragarantia = "SIN FONDOS CONTRAGARANTIA"
    Call Vincula_Tabla(db_catalogos, db_dwh_vf, tbl_sin_fondos_contragarantia, tbl_sin_fondos_contragarantia) ''Vincula con y sin fondos de contragarantía
    
    dbs.Close
    
    'Genera consultas de validación ACOV 201803
    
    Set dbs = OpenDatabase(db_dwh_vf)
    dbs.Execute "Select Producto, SUM(Saldo_Contingente_Mn) AS Saldo_MDP_Suma INTO Valida_NR FROM BD_DWH_NR_" & Mes & " GROUP BY Producto"
    dbs.Execute "Select Producto, SUM(Saldo_Contingente_Mn) AS Saldo_MDP_Suma INTO Valida_R FROM BD_DWH_R_" & Mes & " GROUP BY Producto"
    dbs.Close
    
End Sub




Function Cruza_Catalogos(dbs, T_Final, T_Inicial, T_TipoCambio, T_Programa, T_Udis, T_Agrupamiento, T_TipoCredito, T_TipoGarantia, tbl_sin_fondos_contragarantia)
  dbs.Execute "SELECT A.*, B.TC, C.AGRUPAMIENTO_ID, D.AGRUPAMIENTO, C.ESQUEMA, C.SUBESQUEMA, E.[Paridad_Peso] as CAMBIO, IIF(A.[Monto _Credito_Mn]<900000*E.[Paridad_Peso],0,1) AS [MM_UDIS], F.NR_R, G.CSG " _
    & "INTO [" & T_Final & "_v1] " _
    & "FROM ((((([" & T_Inicial & "] as A LEFT JOIN [" & T_TipoCambio & "] as B " _
    & "ON (YEAR(A.[Fecha_Consulta]) = B.[Año] and MONTH(A.[Fecha_Consulta]) = B.[Mes])) LEFT JOIN [" & T_Programa & "] as C " _
    & "ON (A.[TPRO_CLAVE]=C.[PROGRAMA_ID])) LEFT JOIN [" & T_Agrupamiento & "] as D " _
    & "ON (C.[AGRUPAMIENTO_ID]=D.[AGRUPAMIENTO_ID])) LEFT JOIN [" & T_Udis & "] as E " _
    & "ON (A.[Fecha de Apertura]=E.[Fecha_Paridad])) LEFT JOIN [" & T_TipoCredito & "] as F " _
    & "ON (A.[Tipo_Credito_Id]=F.[Tipo_Credito_ID])) LEFT JOIN [" & T_TipoGarantia & "] as G " _
    & "ON (A.[Tipo_Garantia_Id]=G.[Tipo_garantia_ID]) WHERE(A.[Producto ID] not in (562340,591140,591280,562350)) ; "
    
    'Campo de CSF
    dbs.Execute "SELECT A.*, B.FONDOS_CONTRAGARANTIA " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Final & "_v1] AS A LEFT JOIN [" & tbl_sin_fondos_contragarantia & "] as B " _
    & "ON cstr(A.INTERMEDIARIO_ID) = B.Intermediario_Id AND A.NUMERO_CREDITO = B.CLAVE_CREDITO  " _
    
    'dbs.Execute "DROP TABLE [" & T_Final & "_v1];  "
  
End Function
'Extrae Nombre de Archivos txt
Function ListaArchivosTXT(T_ArchivosTXT, Ruta)
    Set dbs = CurrentDb
        On Error Resume Next
        Archi = Dir$(Ruta & "*.txt")
        Do While Len(Archi)      'Repetir mientras haya archivos
            Archi = Dir$         'Asignar el siguiente archivo
            dbs.Execute "insert into " & T_ArchivosTXT & " (Archivo) " _
                & "Values ('" & Archi & "') ;"
        Loop
    dbs.Close
End Function
Function BorrarListaArchivosTXT(T_ArchivosTXT, Ruta)
    Set dbs = CurrentDb
        Set rst = dbs.OpenRecordset(T_ArchivosTXT)
            With rst
            .MoveFirst
                Do While Not .EOF
                .Delete
                .MoveNext
                Loop
            End With
        rst.Close
    dbs.Close
End Function
'Extrae TC de CAtalogod de EJVV
''Ejemplo Revolventes
''Sub Revolventes_Completo()
''Ruta = "G:\INFO_NAFIN\GARANTIAS\COHORTES\Revolventes\2011"
''Mes = "201101"
''Tabla1 = "Empresarial_Banorte"
''Tabla2 = "Resto"
''Tabla3 = "Revolventes_Resto_201101"
''db_dwh_nr_r_completa = "Foto_R_" & Mes & "_DWH"
''db_dwh_inter = Ruta & "\Revolventes_" & Mes & ".accdb"
''Crea_Tabla_Agrega_Campos db_dwh_inter, "", " * ", "", Tabla1, db_dwh_nr_r_completa, "", "", """"
''Inserta_Filas db_dwh_inter, "", "*", Tabla2, db_dwh_nr_r_completa
''Inserta_Filas db_dwh_inter, "", "*", Tabla3, db_dwh_nr_r_completa
'''"IN '" & db_dwh_inter & "'"
'''End Sub
'Sub GeneraMGIConsejo_20120518()
''var_nr_r, dbs, FotoFinal, Pagos_VF, Recuperaciones, Filtro
'    var_nr_r = "NR"
'    Mes = "201204"
'    Ruta = "G:\INFO_NAFIN\GARANTIAS\COHORTES\DWH\" & Left(Mes, 4) & "\" & Mes & "\"
'    NewBDDestino = "BD_DWH_" & Mes & ".accdb"
'    DestinoNuevo = "BD_DWH_" & Mes & "_VF_Consejo.accdb"
'    db_dwh_vf = Ruta & DestinoNuevo
'    db_dwh_inter = Ruta & NewBDDestino
'    db_dwh_nr_r_completa = "BD_DWH_" & var_nr_r & "_" & Mes
'    Saldos = "BD_DWH_Saldos_" & Mes
'    bd_dwh_nrdb = "BD_DWH_" & var_nr_r & "_" & Mes
'    TC = 12.9942
'    Dim dbs As DAO.Database
'    Set dbs = OpenDatabase(db_dwh_vf)
'    dbs.Execute "SELECT A.*, IIF(Moneda_ID = 54, " & TC & " * B.[Saldo_Contingente],B.[Saldo_Contingente]) as Saldo_Contingente_Mn, " _
'              & "IIF(Moneda_ID = 54, " & TC & " * A.[Monto_Garantizado],A.[Monto_Garantizado]) as Monto_Garantizado_Mn " _
'              & "INTO [" & bd_dwh_nrdb & "] " _
'              & "FROM [" & db_dwh_nr_r_completa & "] A LEFT JOIN [" & Saldos & "] B " _
'              & "ON (A.Intermediario_Id=B.Intermediario_Id AND A.Numero_Credito=B.Numero_Credito) " _
'              & "in '" & db_dwh_inter & "';"
'    dbs.Close
'End Sub
Function Importar_TXT_a_BDAccess(db_dwh_vf, db_dwh_inter, TablaOrigen, Especificaciones)
    Dim objetAccessO As Access.Application
    Set objetAccessO = New Access.Application
    objetAccessO.OpenCurrentDatabase db_dwh_vf
    objetAccessO.DoCmd.TransferText acImportDelim, Especificaciones, TablaOrigen, db_dwh_inter, -1
    objetAccessO.Quit
    Set objetAccessO = Nothing
End Function
Function Inserta_Columna(dbs, TablaMod, Columna, TP_Columna)
    dbs.Execute "alter table [" & TablaMod & "] " _
    & "add column [" & Columna & "] " & TP_Columna & " ; "
End Function
Function Corrige_Campos(dbs, TablaMod, Columna, Condicion, Filtro)
    
    dbs.Execute "update [" & TablaMod & "] as A " _
    & "set [" & Columna & "] = " & Condicion & " " _
    & " " & Filtro & ";"
End Function
Function Corrige_Campos2(dbs, TablaMod, Columna, Condicion, Filtro)
    dbs.Execute "update [" & TablaMod & "] as A " _
    & "set [" & Columna & "] = '" & Condicion & "' " _
    & " " & Filtro & ";"
End Function
'Function Cruza_Tabla_Agrega_Campos_1(dbs, Orden, Linea_1, Campos_Extra, TablaInicial_A, TablaInicial_B, TablaFinal, Filtro_1, Filtro_2)
'        " Select " & Linea_1 & Campos_Extra & " from [" & TablaInicial_A & "] as A left join [" & TablaInicial_B & "] as B & Filtro_1 &  Filtro_2
'End Function
'Function Cruza_Tabla_Agrega_Campos(dbs, db_dwh_vf, Linea_1, Campos_Extra, TablaInicial_A, TablaInicial_B, TablaFinal, Filtro_1, Filtro_2)
'        dbs.Execute "select " & Linea_1 & " " _
'        & " " & Campos_Extra & " " _
'        & "into [" & TablaFinal & "]  " _
'        & "from " & TablaInicial_A & " as A left join " & TablaInicial_B & " as B  " _
'        & " " & Filtro_1 & " " _
'        & " " & Filtro_2 & " " _
'        & " " & db_dwh_vf & "; "
'End Function
'Function Cruza_Inserta_Filas(dbs, db_dwh_vf, Linea_1, Campos_Extra, TablaInicial_A, TablaInicial_B, TablaFinal, Filtro_1, Filtro_2)
'    dbs.Execute "insert into [" & TablaFinal & "] " _
'        & "select " & Linea_1 & " " _
'        & " " & Campos_Extra & " " _
'        & "from " & TablaInicial_A & " as A left join " & TablaInicial_B & " as B " _
'        & " " & Filtro_1 & " " _
'        & " " & Filtro_2 & " " _
'        & " " & db_dwh_vf & "; "
'End Function
Function Crea_Tabla_Agrega_Campos(dbs, Linea_1, Coma, TablaInicial, TablaFinal, Campos_Extra, db_dwh_vf, Filtro, Agrupado_por)
    dbs.Execute "select " & Linea_1 & Coma & " " _
        & " " & Campos_Extra & " " _
        & " into [" & TablaFinal & "] " _
        & " from " & TablaInicial & " as A " _
        & " " & db_dwh_vf & " " _
        & " " & Filtro & " " _
        & " " & Agrupado_por & "; "
End Function
'Sub Inserta_Filas(dbs, db_dwh_vf, Linea_1, TablaInicial, TablaFinal, Filtro, Agrupado_por)
'    dbs.Execute "insert into [" & TablaFinal & "] " _
''        & "select " & Linea_1 & " " _
'        & "from [" & TablaInicial & "] " _
'        & " " & db_dwh_vf & " " _
'        & " " & Filtro & " " _
'        & " " & Agrupado_por & "; "
'End Sub
Function LinkSchema(RutaBases, RutaDestino, Base, archivo_txt)
   Dim db As DAO.Database, tbl As TableDef
   Set db = OpenDatabase(RutaDestino & Base)
   Set tbl = db.CreateTableDef(Left(archivo_txt, Len(archivo_txt) - 4))

   tbl.Connect = "Text;DATABASE=" & RutaBases & ";TABLE=" & archivo_txt & ""
  
   tbl.SourceTableName = archivo_txt
   db.TableDefs.Append tbl
   db.TableDefs.Refresh
   db.Close
End Function

Function Busca_Agrupamientos_Vacios(dbs, db_dwh_nr_r_completa)
    Dim sql2 As String
    sql2 = "COUNT AGRUPAMIENTO_ID FROM " & [db_dwh_nr_r_completa] & " WHERE AGRUPAMIENTO_ID IS NULL"
    DoCmd.RunSQL sql2
    
End Function
