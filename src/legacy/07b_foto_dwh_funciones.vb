Function Inserta_Columna(dbs, TablaMod, Columna, TP_Columna)
    dbs.Execute "alter table [" & TablaMod & "] " _
    & "add column [" & Columna & "] " & TP_Columna & " ; "
End Function

Function Corrige_Campos(dbs, TablaMod, Columna, Condicion, Filtro)
    
    dbs.Execute "update [" & TablaMod & "] as A " _
    & "set [" & Columna & "] = " & Condicion & " " _
    & " " & Fil
End Function

Function Cruza_Catalogos(dbs, T_Final, T_Inicial, T_TipoCambio, T_Programa, T_Udis, T_Agrupamiento, T_TipoCredito, T_TipoGarantia)
dbs.Execute "SELECT A.*, B.TC, C.AGRUPAMIENTO_ID, D.AGRUPAMIENTO, C.ESQUEMA, C.SUBESQUEMA, E.[Paridad_Peso] as CAMBIO, IIF(A.[Monto _Credito_Mn]<900000*E.[Paridad_Peso],0,1) AS [MM_UDIS], F.NR_R, G.CSG " _
    & "INTO [" & T_Final & "] " _
    & "FROM ((((([" & T_Inicial & "] as A LEFT JOIN [" & T_TipoCambio & "] as B " _
    & "ON (YEAR(A.[Fecha_Consulta]) = B.[Año] and MONTH(A.[Fecha_Consulta]) = B.[Mes])) LEFT JOIN [" & T_Programa & "] as C " _
    & "ON (A.[TPRO_CLAVE]=C.[PROGRAMA_ID])) LEFT JOIN [" & T_Agrupamiento & "] as D " _
    & "ON (C.[AGRUPAMIENTO_ID]=D.[AGRUPAMIENTO_ID])) LEFT JOIN [" & T_Udis & "] as E " _
    & "ON (A.[Fecha de Apertura]=E.[Fecha_Paridad])) LEFT JOIN [" & T_TipoCredito & "] as F " _
    & "ON (A.[Tipo_Credito_Id]=F.[Tipo_Credito_ID])) LEFT JOIN [" & T_TipoGarantia & "] as G " _
    & "ON (A.[Tipo_Garantia_Id]=G.[Tipo_garantia_ID]) WHERE(A.[Producto ID] not in (562340,591140,591280,562350)) ; "
    'Call Corrige_Campos(dbs, T_Final, "TC", 1, " Where (A.Moneda_Id=1) ")
End Function

Function Crea_Tabla_Agrega_Campos(dbs, Linea_1, Coma, TablaInicial, TablaFinal, Campos_Extra, BaseDestino, Filtro, Agrupado_por)
    dbs.Execute "select " & Linea_1 & Coma & " " _
        & " " & Campos_Extra & " " _
        & " into [" & TablaFinal & "] " _
        & " from " & TablaInicial & " as A " _
        & " " & BaseDestino & " " _
        & " " & Filtro & " " _
        & " " & Agrupado_por & "; "
End Function

Function Compara_Registros(BaseOrigen, BaseOrigen_2, Tabla_Contar_Registros_1, Tabla_Contar_Registros_2, Campo_TCR1, Campo_TCR2) As String
    No_Registros_1 = Cuenta_Registros(BaseOrigen, Tabla_Contar_Registros_1, Campo_TCR1)
    No_Registros_2 = Cuenta_Registros(BaseOrigen_2, Tabla_Contar_Registros_2, Campo_TCR2)
    If No_Registros_1 = No_Registros_2 Then
        Compara_Registros = "Listo"     'MsgBox "Tabla Tipo de Persona Depurada Lista"
    ElseIf No_Registros_1 < No_Registros_2 Then
        MsgBox "Existen menos registros en la Tabla " & Tabla_Contar_Registros_1 & " que en la tabla " & Tabla_Contar_Registros_2 & ". Favor de Validar"
        Compara_Registros = No_Registros_1 & "-" & Tabla_Contar_Registros_1 & "," & Tabla_Contar_Registros_2 & "-" & No_Registros_2
    ElseIf No_Registros_1 > No_Registros_2 Then
        MsgBox "Existen menos registros en la Tabla " & Tabla_Contar_Registros_2 & " que en la tabla " & Tabla_Contar_Registros_1 & ". Favor de Validar"
        Compara_Registros = No_Registros_1 & "-" & Tabla_Contar_Registros_1 & "," & Tabla_Contar_Registros_2 & "-" & No_Registros_2        'MsgBox "El Número de registro de la Tabla sin Repetidos es mayor al Número de registros de la Tabla Depurada. Favor de Validar"
    End If
End Function

Function Cuenta_Registros(BaseOrigen, Tabla_Contar_Registros_1, Campo_TCR1) As Double
    Dim BDD As Database
    Dim tbl As Recordset
    Dim SQL As String
    Set BDD = OpenDatabase(BaseOrigen)
        SQL = "Select COUNT(A." & Campo_TCR1 & ") from (Select " & Campo_TCR1 & " from [" & Tabla_Contar_Registros_1 & "]) as A;"
    Set tbl = BDD.OpenRecordset(SQL)
         Cuenta_Registros = tbl("expr1000")
    tbl.Close
    BDD.Close
End Function

