Function agrupa_pagos(dbs, base_origen, tbl_final, tbl_inicial)
dbs.Execute "SELECT Producto, Intermediario_Id, Numero_Credito, Fecha_Garantia_Honrada, [MM_UDIS], NR_R, " _
          & "[Razón Social (Intermediario)], TPRO_CLAVE, AGRUPAMIENTO,  " _
          & "sum(Monto_Desembolso_Mn) as Monto_Desem_Mn, " _
          & "sum(Interes_Desembolso_Mn) as Interes_Desem_Mn, " _
          & "sum([Interes_Moratorios_Mn]) as Interes_Morat_Mn, " _
          & "sum(Monto_Pagado_Mn) as MPagado_Mn " _
          & "INTO [" & tbl_final & "] " _
          & "FROM [" & tbl_inicial & "] " _
          & "IN '" & base_origen & "' " _
          & "GROUP BY Producto, Intermediario_Id, Numero_Credito, Fecha_Garantia_Honrada, [MM_UDIS], NR_R, [Razón Social (Intermediario)], TPRO_CLAVE, AGRUPAMIENTO; "
End Function

Function agrupa_recuperaciones(dbs, tbl_final, tbl_inicial)
dbs.Execute "SELECT Estatus, Intermediario_Id, Numero_Credito, Concatenar_Saldos, " _
          & "sum(Monto_Mn) as Monto_Recup_Mn, " _
          & "sum(Interes_Mn) as Interes_Recup_Mn, " _
          & "sum(Moratorios_Mn) as Moratorios_Recup_Mn, " _
          & "sum(Excedente_Mn) as Excedente_Recup_Mn, " _
          & "sum(Monto_Total_Mn) as Monto_Total_Recup_Mn " _
          & "INTO [" & tbl_final & "] " _
          & "FROM [" & tbl_inicial & "] " _
          & "GROUP BY Estatus, Intermediario_Id, Numero_Credito, Concatenar_Saldos; "
End Function

Function campos_calculados_recuperaciones(dbs, tbl_recuperaciones_f, tbl_recuperaciones_i, tbl_estatus)
dbs.Execute "SELECT A.*, " _
          & "A.Monto*A.TC as Monto_Mn, " _
          & "A.Interes*A.TC as Interes_Mn, " _
          & "A.Moratorios*A.TC as Moratorios_Mn, " _
          & "A.Excedente*A.TC as Excedente_Mn, " _
          & "A.[Gastos Juicio]*A.TC as Gasto_Juicio_Mn, " _
          & "(IIF(A.Monto is null,0,A.Monto)+IIF(A.Interes is null,0,A.Interes)+IIF(A.Moratorios is null,0,A.Moratorios)+IIF(A.Excedente is null,0,A.Excedente))*A.TC as Sub_Total_Mn, " _
          & "(IIF(A.Monto is null,0,A.Monto)+IIF(A.Interes is null,0,A.Interes)+IIF(A.Moratorios is null,0,A.Moratorios)+IIF(A.Excedente is null,0,A.Excedente)-IIF(A.[Gastos Juicio] is null,0,A.[Gastos Juicio]))*A.TC as Monto_Total_Mn, " _
          & "B.[Recup/Rescat] as Recup_Rescat " _
          & "INTO [" & tbl_recuperaciones_f & "] " _
          & "FROM [" & tbl_recuperaciones_i & "] as A left join [" & tbl_estatus & "] as B " _
          & "ON A.[Estatus]=B.[Estatus ID]; "
End Function

Function compacta_repara(database)
    database_ = Left(database, Len(database) - 6)
    compact_database = database_ & "_temp" & ".accdb"
    DAO.DBEngine.CompactDatabase database, compact_database
    Kill database
    Name compact_database As database
End Function

Function copiar_tabla_bd(wd_origen, wd_destino, tbl_origen, tbl_destino)
     Dim objetAccessO As Access.Application
     Dim objetAccessD As Access.Application
     Set objetAccessO = New Access.Application
     Set objetAccessD = New Access.Application
     objetAccessO.OpenCurrentDatabase wd_origen
     objetAccessO.DoCmd.CopyObject wd_destino, tbl_destino, acTable, tbl_origen
     objetAccessO.Quit
     objetAccessD.Quit
     Set objetAccessO = Nothing
     Set objetAccessD = Nothing
End Function

Function corrige_campos(dbs, tbl_mod, columna, condicion, filtro)
    dbs.Execute "update [" & tbl_mod & "] as A " _
    & "set [" & columna & "] = " & condicion & " " _
    & " " & filtro & ";"
End Function

Function corrige_fecha_pago(dbs, tabla)
dbs.Execute "UPDATE " & tabla & " " _
          & "SET Fecha_Garantia_Honrada=#02/08/2011# " _
          & "WHERE Numero_Credito='9842725312' and Intermediario_Id='10040012'; "
End Function

Function cruza_catalogos_2(dbs, tbl_final, tbl_inicial, tbl_tipo_cambio, tbl_programa, tbl_udis, tbl_agrupamiento, tbl_tipo_credito, tbl_tipo_garantia, parametro_dif, tbl_sin_fondos)
'si sale un error es por el tipo de datos del intermediario id osea hay q po ner un cstr()
dbs.Execute "SELECT A.*, B.TC, C.AGRUPAMIENTO_ID, D.AGRUPAMIENTO, C.ESQUEMA, C.SUBESQUEMA, E.[Paridad_Peso] as CAMBIO, IIF(A.[Monto _Credito_Mn]<900000*E.[Paridad_Peso],0,1) AS [MM_UDIS], F.NR_R, G.CSG, IIF(H.FONDOS_CONTRAGARANTIA='SF','SF','CF') as CSF  " _
    & "INTO [" & tbl_final & "] " _
    & "FROM (((((([" & tbl_inicial & "] as A LEFT JOIN [" & tbl_tipo_cambio & "] as B " _
    & "ON (YEAR(A.[Fecha_Consulta]) = B.[Año] and MONTH(A.[Fecha_Consulta]) = B.[Mes])) LEFT JOIN [" & tbl_programa & "] as C " _
    & "ON (A.[TPRO_CLAVE]=C.[PROGRAMA_ID])) LEFT JOIN [" & tbl_agrupamiento & "] as D " _
    & "ON (C.[AGRUPAMIENTO_ID]=D.[AGRUPAMIENTO_ID])) LEFT JOIN [" & tbl_udis & "] as E " _
    & "ON (A.[" & parametro_dif & "]=E.[Fecha_Paridad])) LEFT JOIN [" & tbl_tipo_credito & "] as F " _
    & "ON (A.[Tipo_Credito_Id]=F.[Tipo_Credito_ID])) LEFT JOIN [" & tbl_tipo_garantia & "] as G " _
    & "ON (A.[Tipo_Garantia_Id]=G.[Tipo_garantia_ID])) LEFT JOIN [" & tbl_sin_fondos & "] as H " _
    & "ON (cstr(A.[Intermediario_Id])=H.[Intermediario_Id] AND A.[Numero_Credito]=H.[CLAVE_CREDITO]);"
    Call corrige_campos(dbs, tbl_final, "TC", 1, " Where (A.Moneda_Id=1) ")
End Function

Function cruza_pagof1_pagof2(base_pagos, tbl_pagos_final_inter, tbl_pagos_f1, tbl_pagos_f2)
base_pagos.Execute "SELECT B.Concatenado_P2 as Concatenado, B.Fecha_Consulta, B.Intermediario_Id, B.Numero_Credito, A.Producto, B.[Pago ID], " _
          & "A.[Razón Social (Intermediario)], LEFT(B.[MIN Fecha_Registro],10) as MIN_Fecha_Registro, A.Fecha_Garantia_Honrada, " _
          & "A.TPRO_CLAVE, A.Programa_Original, A.Programa_Id, A.[Monto _Credito_Mn], A.Moneda_Id, " _
          & "A.[Fecha de Apertura], A.Tipo_Garantia_Id, A.TIPO_PERSONA, A.[RFC Empresa / Acreditado], " _
          & "B.[SUM Monto_Desembolso] AS Monto_Desembolsado, " _
          & "B.[SUM Interes_Desembolso] AS Interes_Desembolso, " _
          & "B.[SUM Intereses Moratorios] AS Interes_Moratorios, " _
          & "A.Porcentaje_Garantizado, A.Tipo_Credito_Id, A.Estatus_Recuperacion, " _
          & "A.[Empresa / Acreditado (Descripción)], A.[Fecha Registro Alta] " _
          & "INTO [" & tbl_pagos_final_inter & "] " _
          & "FROM " & tbl_pagos_f2 & " as B LEFT JOIN " & tbl_pagos_f1 & " as A " _
          & "ON A.Concatenado_P1 = B.Concatenado_P2; "
End Function

 "SELECT B.Concatenado_P2 as Concatenado, B.Fecha_Consulta, B.Intermediario_Id, B.Numero_Credito, A.Producto, B.[Pago ID],
        A.[Razón Social (Intermediario)], LEFT(B.[MIN Fecha_Registro],10) as MIN_Fecha_Registro, A.Fecha_Garantia_Honrada,
        A.TPRO_CLAVE, A.Programa_Original, A.Programa_Id, A.[Monto _Credito_Mn], A.Moneda_Id,
        A.[Fecha de Apertura], A.Tipo_Garantia_Id, A.TIPO_PERSONA, A.[RFC Empresa / Acreditado],
        B.[SUM Monto_Desembolso] AS Monto_Desembolsado,
        B.[SUM Interes_Desembolso] AS Interes_Desembolso,
        B.[SUM Intereses Moratorios] AS Interes_Moratorios,
        A.Porcentaje_Garantizado, A.Tipo_Credito_Id, A.Estatus_Recuperacion,
        A.[Empresa / Acreditado (Descripción)], A.[Fecha Registro Alta]
        INTO [Pagadas_Global_VF_202412_Inter]
        FROM DWH_Pagos_F2_202412 as B LEFT JOIN DWH_Pagos_F1_202412 as A
        ON A.Concatenado_P1 = B.Concatenado_P2;"

Function cuadro_valida_recup_dwh_dac(dbs, tbl_recuperaciones_f, tbl_recuperaciones_i)
    dbs.Execute "SELECT Historico, Tipo_Cambio_Cierre, Producto, " _
          & "Sum(Monto) AS S_Monto, Sum(Interes) AS S_Interes, " _
          & "Sum(Moratorios) AS S_Moratorios, Sum(Excedente) AS S_Excedente, " _
          & "Sum([Gastos Juicio]) AS S_Gastos_Juicio " _
          & "INTO [" & tbl_recuperaciones_f & "] " _
          & "FROM [" & tbl_recuperaciones_i & "] " _
          & "GROUP BY Historico, Tipo_Cambio_Cierre, Producto; "
End Function

Function cuadro_valida_recup_td(dbs, tbl_recuperaciones_f, tbl_recuperaciones_i)
    dbs.Execute "SELECT  Producto, Tipo_Cambio_Cierre, " _
          & "Sum(Monto) AS S_Monto, Sum(Interes) AS S_Interes, " _
          & "Sum(Moratorios) AS S_Moratorios, Sum(Excedente) AS S_Excedente, " _
          & "Sum([Gastos Juicio]) AS S_Gastos_Juicio " _
          & "INTO [" & tbl_recuperaciones_f & "] " _
          & "FROM [" & tbl_recuperaciones_i & "] " _
          & "GROUP BY Producto, Tipo_Cambio_Cierre; "
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

Function exporta_access_excel(base_destino, base_origen, tbl_origen)
    Dim objetAccessO As Access.Application
    Set objetAccessO = New Access.Application
    objetAccessO.OpenCurrentDatabase base_origen
    objetAccessO.DoCmd.TransferSpreadsheet 1, 10, tbl_origen, base_destino, True
    objetAccessO.Quit
    Set objetAccessO = Nothing
End Function

Function extracto_bd(dbs, db_inicial, db_extracto, condicion)
dbs.Execute "SELECT * INTO " & db_extracto & " FROM " & db_inicial & " " _
           & condicion & "; "
End Function

Function importa_recuperaciones(dbs, base_origen, tbl_final, tbl_inicial)
dbs.Execute "SELECT *, Intermediario_Id & Numero_Credito as Concatenar_Saldos " _
          & "INTO [" & tbl_final & "] " _
          & "FROM [" & tbl_inicial & "] " _
          & "IN '" & base_origen & "'; "
End Function

Function inserta_columna(dbs, tbl_mod, columna, tp_columna)
    dbs.Execute "alter table [" & tbl_mod & "] " _
    & "add column [" & columna & "] " & tp_columna & " ; "
End Function

Function inserta_filas_in(dbs, base_origen, tbl_inicial, tbl_final)
    dbs.Execute "insert into [" & tbl_final & "] " _
        & "select * from [" & tbl_inicial & "]" _
        & base_origen & ";"
End Function

Function mes_letra(mes) As String
    Select Case mes
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
 

Sub ordena_tabla_final(dbs, tbl_final, tbl_inicial)
dbs.Execute "SELECT Concatenar_Saldos, Fecha_Consulta, Programa_Id, " _
          & "Tipo_Garantia_Id, Tipo_Cambio_Cierre, Programa_Original, Porcentaje_Garantizado, " _
          & "[Monto _Credito_Mn], Fecha_Apertura, Moneda_Id, Tipo_Credito_Id, Fecha_Garantia_Honrada, TPRO_CLAVE, " _
          & "NR_R, Producto, AGRUPAMIENTO_ID, AGRUPAMIENTO, Intermediario_Id, " _
          & "Numero_Credito, Id, Monto, Interes, Moratorios, Descripcion, Estatus, Fecha_Registro, Fecha, Monto_Mn, " _
          & "Interes_Mn, Excedente, Excedente_Mn, Moratorios_Mn, Sub_Total_Mn, [Razón Social (Intermediario)], " _
          & "[Empresa / Acreditado (Descripción)], [RFC Empresa / Acreditado], TIPO_PERSONA, Recup_Rescat, " _
          & "MM_UDIS, ESQUEMA, CSG, CSF, Gasto_Juicio_Mn as Gastos_Juicio_Mn, Monto_Total_Mn, MPagado_Mn, " _
          & "Historico, [Fecha Registro Alta] " _
          & "INTO [" & tbl_final & "] " _
          & "FROM [" & tbl_inicial & "]; "
End Sub

Function unir_agrupados_rfc(dbs, tbl_final, tbl_left, tbl_right)
dbs.Execute "SELECT A.*, B.Estatus, B.Monto_Recup_Mn, B.Interes_Recup_Mn, " _
          & "B.Moratorios_Recup_Mn, B.Excedente_Recup_Mn, B.Monto_Total_Recup_Mn " _
          & "INTO " & tbl_final & " " _
          & " FROM [" & tbl_left & "] A LEFT JOIN [" & tbl_right & "] B " _
          & "ON A.Concatenar_Saldos=B.Concatenar_Saldos; "
End Function

Function unir_flujos_p_u_r(dbs, tbl_final, tbl_left, tbl_right)
dbs.Execute "SELECT A.*, B.Monto_Mn as Monto_Recup_Mn, B.Interes_Mn as Interes_Recup_Mn, " _
          & "B.Moratorios_Mn as Moratorios_Recup_Mn, B.Excedente_Mn as Excedentes_Recup_Mn, B.Monto_Total_Mn as Monto_Total_Recup_Mn " _
          & "INTO [" & tbl_final & "] " _
          & "FROM [" & tbl_left & "] A LEFT JOIN [" & tbl_right & "] B " _
          & "ON A.Concatenar_Saldos=B.Concatenar_Saldos; "
End Function

Function unir_flujos_r_u_p(dbs, tbl_final, tbl_left, tbl_right)
dbs.Execute "SELECT A.*, B.MPagado_Mn " _
          & "INTO [" & tbl_final & "] " _
          & "FROM [" & tbl_left & "] A LEFT JOIN [" & tbl_right & "] B " _
          & "ON A.Concatenar_Saldos=B.Concatenar_Saldos; "
End Function

Function vincula_tabla(base_pagos, base_destino, tbl_origen, tbl_destino)
    Dim objetAccessO As Access.Application
    Set objetAccessO = New Access.Application
    objetAccessO.OpenCurrentDatabase base_destino
    objetAccessO.DoCmd.TransferDatabase acLink, "Microsoft Access", base_pagos, acTable, tbl_origen, tbl_destino
    objetAccessO.Quit
    Set objetAccessO = Nothing
End Function
