Sub Principal_TP()
Dim dbs As DAO.database

Mes = "202411"
Mes_Anterior = "202410"

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

db_foto_revolventes_cohortes_preliminar_anterior = wd_external & "FotoRevolventesCohortes_" & Mes_Anterior & "_Preeliminar.accdb"
db_foto_revolventes_cohortes_preliminar = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & Mes & "_Preeliminar.accdb"

tbl_vf_foto_r = "VF_Foto_R_" & Mes
tbl_pfpm = "PFPM_" & Mes
tbl_repetidos_tp = "RepetidosTP"
tbl_repetidos_tp_base = "Repetidos_TP_Base"
tbl_vf_pfpm = "VF_PFPM_" & Mes
tbl_repetidos_tp_concentrado = "RepetidosTP_Concentrado"
tbl_repetidos_tp_concentrado_depurado = "RepetidosTP_Concentrado_Depurado"
tbl_repetidos_nuevos = "RepetidosNuevos"

linea_1 = "A.BANCO, A.NOMBRE, A.TAXONOMIA, A.AGRUPAMIENTO, A.TIPO_PERSONA "
linea_2 = "A.BANCO, A.NOMBRE, A.TAXONOMIA, A.AGRUPAMIENTO "
filtro_0 = "ON (A.BANCO=B.BANCO and A.NOMBRE=B.NOMBRE and A.TAXONOMIA=B.TAXONOMIA and A.AGRUPAMIENTO=B.AGRUPAMIENTO and A.TIPO_PERSONA=B.TIPO_PERSONA)"
filtro_1 = "ON (A.BANCO=B.BANCO and A.NOMBRE=B.NOMBRE and A.TAXONOMIA=B.TAXONOMIA and A.AGRUPAMIENTO=B.AGRUPAMIENTO)"
filtro_null_0 = "A.BANCO_2 is Null"
filtro_null_1 = "( A.BANCO_3 is Null or A.TIPO_PERSONA=A.BANCO_3 )"
filtro_final = "ON (A.BANCO & A.NOMBRE & A.TAXONOMIA & A.AGRUPAMIENTO & A.TIPO_PERSONA<>B.BANCO_3)"

'Copia la tbl_repetidos_tp_base del archivo del mes anterior

If ExisteTabla(db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_base) = False Then
    CopiarTabla_BD db_foto_revolventes_cohortes_preliminar_anterior, db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_base, tbl_repetidos_tp_base
Else
    MsgBox "Ya no se copio la tabla: " & tbl_repetidos_tp_base & " porque ya existe"
    
End If

Genera_Tipo_Persona db_foto_revolventes_cohortes_preliminar, linea_1, tbl_pfpm, tbl_vf_foto_r, ""        'Genera Katalogo con Repetidos

RevisaRepetidos db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp, tbl_pfpm

'ACOV 201906 aquí se insertan los faltantes
Marca_RegistrossRepetir db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_concentrado, tbl_repetidos_tp, tbl_repetidos_tp_base, filtro_0, "BANCO_2"
Bandera1 = Compara_Registros(db_foto_revolventes_cohortes_preliminar, db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_concentrado, tbl_repetidos_tp_base, "BANCO_2", "BANCO")
LimpiaNull db_foto_revolventes_cohortes_preliminar, "Temp_" & tbl_repetidos_tp_concentrado_depurado, tbl_repetidos_tp_concentrado, filtro_null_0
If Bandera1 <> "Listo" Then
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_pfpm
    Exit Sub
End If
Marca_RegistrossRepetir_2 db_foto_revolventes_cohortes_preliminar, "Temp2_" & tbl_repetidos_tp_concentrado_depurado, "Temp_" & tbl_repetidos_tp_concentrado_depurado, tbl_repetidos_tp_base, filtro_1, "BANCO_3"
Bandera2 = Compara_Registros(db_foto_revolventes_cohortes_preliminar, db_foto_revolventes_cohortes_preliminar, "Temp2_" & tbl_repetidos_tp_concentrado_depurado, tbl_repetidos_tp_base, "BANCO_3", "BANCO")
LimpiaNull db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_concentrado_depurado, "Temp2_" & tbl_repetidos_tp_concentrado_depurado, filtro_null_1
If Bandera2 <> "Listo" Then
    BorraTabla db_foto_revolventes_cohortes_preliminar, "Temp_" & tbl_repetidos_tp_concentrado_depurado
    BorraTabla db_foto_revolventes_cohortes_preliminar, "Temp2_" & tbl_repetidos_tp_concentrado_depurado
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_pfpm
    Exit Sub
End If
No_Registros_1 = Cuenta_Registros(db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_concentrado_depurado, "BANCO")

    BorraTabla db_foto_revolventes_cohortes_preliminar, "Temp_" & tbl_repetidos_tp_concentrado_depurado
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_concentrado
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_concentrado_depurado

If No_Registros_1 = 0 Then
    Genera_TP_VF db_foto_revolventes_cohortes_preliminar, tbl_vf_pfpm, tbl_pfpm, "Temp2_" & tbl_repetidos_tp_concentrado_depurado
    BorraTabla db_foto_revolventes_cohortes_preliminar, tbl_pfpm
    BorraTabla db_foto_revolventes_cohortes_preliminar, "Temp2_" & tbl_repetidos_tp_concentrado_depurado
    MsgBox "Los registros repetidos son los mismos que existen en la BD: " & tbl_repetidos_tp_base & ". Lista la BD de Tipo de Persona"
Else
    MsgBox "Existen registros repetidos diferentes a los de la BD: " & tbl_repetidos_tp_base & ". " _
    & "Estos registro se encuentran en la BD: " & tbl_repetidos_tp_concentrado_depurado & " " _
    & "Favor de seleccionar los registros correctos y darlos de alta en la BD: " & tbl_repetidos_tp_base & " " _
    & ". Posteriormente volver a correr el procesode Tipo de Persona"
End If

'BorraTabla
'CopiarTabla_BD db_foto_revolventes_cohortes_preliminar, RutaBaseDatosDestino, tbl_vf_pfpm, tbl_vf_pfpm
End Sub


Function Genera_Tipo_Persona(db_foto_revolventes_cohortes_preliminar, linea_1, tbl_pfpm, tbl_vf_foto_r, Condicion1)
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "SELECT " & linea_1 & " " _
          & "INTO " & tbl_pfpm & " " _
          & "FROM [" & tbl_vf_foto_r & "] AS A " _
          & "IN '" & db_foto_revolventes_cohortes_preliminar & "' " _
          & " " & Condicion1 & " " _
          & "GROUP BY  " & linea_1 & " " _
          & "ORDER BY " & linea_1 & "; "
    dbs.Close
End Function
Function Marca_RegistrossRepetir(db_foto_revolventes_cohortes_preliminar, tbl_vf_pfpm, tbl_pfpm, TPsRepetidos, Filtro, CampoExtra)
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "SELECT A.*, " _
          & "B.BANCO & B.NOMBRE & B.TAXONOMIA & B.AGRUPAMIENTO& B.TIPO_PERSONA as " & CampoExtra & " " _
          & "INTO [" & tbl_vf_pfpm & "] " _
          & "FROM [" & tbl_pfpm & "] A LEFT JOIN [" & TPsRepetidos & "] B " _
          & " " & Filtro & " " _
          & "IN '" & db_foto_revolventes_cohortes_preliminar & "'; "
    dbs.Close
End Function
Function Marca_RegistrossRepetir_2(db_foto_revolventes_cohortes_preliminar, tbl_vf_pfpm, tbl_pfpm, TPsRepetidos, Filtro, CampoExtra)
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "SELECT A.*, " _
          & "B.BANCO & B.NOMBRE & B.TAXONOMIA & B.AGRUPAMIENTO& A.TIPO_PERSONA as " & CampoExtra & " " _
          & "INTO [" & tbl_vf_pfpm & "] " _
          & "FROM [" & tbl_pfpm & "] A LEFT JOIN [" & TPsRepetidos & "] B " _
          & " " & Filtro & " " _
          & "IN '" & db_foto_revolventes_cohortes_preliminar & "'; "
    dbs.Close
End Function
Function LimpiaNull(db_foto_revolventes_cohortes_preliminar, tbl_vf_pfpm, tbl_pfpm, Filtro)
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "SELECT A.* " _
          & "INTO [" & tbl_vf_pfpm & "] " _
          & "FROM [" & tbl_pfpm & "] as A  " _
          & "IN '" & db_foto_revolventes_cohortes_preliminar & "' " _
          & "WHERE " & Filtro & "  ; "
End Function
Function BorraTabla(db_foto_revolventes_cohortes_preliminar, Tabla)
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "drop table " & Tabla & ";"
    dbs.Close
End Function
Function RevisaRepetidos(db_foto_revolventes_cohortes_preliminar, RepetidosTP, tbl_pfpm)
    BANCO_1 = "Inicio"
    Bandera = 0
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    Set rst = dbs.OpenRecordset(tbl_pfpm)
    With rst
        .MoveFirst
        Do While Not .EOF
            If BANCO_1 = "Inicio" Then
                Banco1 = .BANCO
                NOMBRE1 = .NOMBRE
                TAXONOMIA1 = .TAXONOMIA
                AGRUPAMIENTO1 = .AGRUPAMIENTO
                TIPO_PERSONA1 = .TIPO_PERSONA
                BANCO_1 = .BANCO
                NOMBRE_1 = .NOMBRE
                TAXONOMIA_1 = .TAXONOMIA
                AGRUPAMIENTO_1 = .AGRUPAMIENTO
                TIPO_PERSONA_1 = .TIPO_PERSONA
            Else
                Banco1 = .BANCO
                NOMBRE1 = .NOMBRE
                TAXONOMIA1 = .TAXONOMIA
                AGRUPAMIENTO1 = .AGRUPAMIENTO
                TIPO_PERSONA1 = .TIPO_PERSONA
                If Banco1 = BANCO_1 And NOMBRE_1 = NOMBRE1 And TAXONOMIA_1 = TAXONOMIA1 And AGRUPAMIENTO_1 = AGRUPAMIENTO1 Then
                    If Bandera = 0 Then
                        dbs.Execute "Select  '" & Banco1 & "' as BANCO, '" & NOMBRE1 & "' as NOMBRE, " _
                        & "'" & TAXONOMIA1 & "' as TAXONOMIA, '" & AGRUPAMIENTO1 & "' as AGRUPAMIENTO, '" & TIPO_PERSONA1 & "' as TIPO_PERSONA " _
                        & "into [" & RepetidosTP & "];"
                        dbs.Execute "insert into [" & RepetidosTP & "] " _
                        & "(BANCO, NOMBRE, TAXONOMIA, AGRUPAMIENTO, TIPO_PERSONA) values " _
                        & "('" & BANCO_1 & "', '" & NOMBRE_1 & "', '" & TAXONOMIA_1 & "', '" & AGRUPAMIENTO_1 & "', '" & TIPO_PERSONA_1 & "');"
                        Bandera = 1
                    Else
                        dbs.Execute "insert into [" & RepetidosTP & "]  " _
                        & "(BANCO, NOMBRE, TAXONOMIA, AGRUPAMIENTO, TIPO_PERSONA) values " _
                        & "('" & Banco1 & "', '" & NOMBRE1 & "','" & TAXONOMIA1 & "', '" & AGRUPAMIENTO1 & "', '" & TIPO_PERSONA1 & "');"
                        dbs.Execute "insert into [" & RepetidosTP & "]  " _
                        & "(BANCO, NOMBRE, TAXONOMIA, AGRUPAMIENTO, TIPO_PERSONA) values " _
                        & "('" & BANCO_1 & "', '" & NOMBRE_1 & "','" & TAXONOMIA_1 & "', '" & AGRUPAMIENTO_1 & "', '" & TIPO_PERSONA_1 & "');"
                    End If
                Else
                    BANCO_1 = Banco1
                    NOMBRE_1 = NOMBRE1
                    TAXONOMIA_1 = TAXONOMIA1
                    AGRUPAMIENTO_1 = AGRUPAMIENTO1
                    TIPO_PERSONA_1 = TIPO_PERSONA1
                End If
            End If
        .MoveNext
        Loop
    End With
    rst.Close
    dbs.Close
End Function
Function Cuenta_Registros(db_foto_revolventes_cohortes_preliminar, Tabla_Contar_Registros_1, Campo_TCR1) As Double
    Dim BDD As database
    Dim tbl As Recordset
    Dim SQL As String
    Set BDD = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
        SQL = "Select COUNT(A." & Campo_TCR1 & ") from (Select " & Campo_TCR1 & " from [" & Tabla_Contar_Registros_1 & "]) as A;"
    Set tbl = BDD.OpenRecordset(SQL)
         Cuenta_Registros = tbl("expr1000")
    tbl.Close
    BDD.Close
End Function
Function Compara_Registros(db_foto_revolventes_cohortes_preliminar, db_foto_revolventes_cohortes_preliminar_2, Tabla_Contar_Registros_1, Tabla_Contar_Registros_2, Campo_TCR1, Campo_TCR2) As String
    No_Registros_1 = Cuenta_Registros(db_foto_revolventes_cohortes_preliminar, Tabla_Contar_Registros_1, Campo_TCR1)
    No_Registros_2 = Cuenta_Registros(db_foto_revolventes_cohortes_preliminar_2, Tabla_Contar_Registros_2, Campo_TCR2)
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
Function Union_bases(data_base, BaseInicial, BaseFinal, Linea3)
    Dim dbs As database
    Set dbs = OpenDatabase(data_base)
    dbs.Execute "INSERT INTO [" & BaseFinal & "] SELECT * FROM [" & BaseInicial & "] " _
    & "ORDER BY " & Linea3 & ";"        'Revisar porque no ordena
    dbs.Close
End Function
Function Primer_base(data_base, BaseInicial, BaseFinal)
    Dim dbs As database
    Set dbs = OpenDatabase(data_base)
    dbs.Execute "SELECT * " _
    & "INTO " & BaseFinal & " " _
    & "FROM " & BaseInicial & ";"
    dbs.Close
End Function
Function Genera_TP_VF(db_foto_revolventes_cohortes_preliminar, VF_TP, tbl_pfpm, TP_Repetidos_VF)
    Bandera_Inicio = 0
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    Set rst = dbs.OpenRecordset(tbl_pfpm)
    With rst
        .MoveFirst
        Do While Not .EOF
                Banco1 = .BANCO
                NOMBRE1 = .NOMBRE
                TAXONOMIA1 = .TAXONOMIA
                AGRUPAMIENTO1 = .AGRUPAMIENTO
                TIPO_PERSONA1 = .TIPO_PERSONA
                Bandera = 0
                Set rst2 = dbs.OpenRecordset(TP_Repetidos_VF)
                With rst2
                .MoveFirst
                Do While Not .EOF
                    BANCO_3 = .BANCO_3
                    If Banco1 & NOMBRE1 & TAXONOMIA1 & AGRUPAMIENTO1 & TIPO_PERSONA1 = BANCO_3 And Bandera = 0 Then
                        Bandera = 1
                    End If
                .MoveNext
                Loop
                End With
        If Bandera = 0 And Bandera_Inicio = 0 Then
            dbs.Execute "Select  '" & Banco1 & "' as BANCO, '" & NOMBRE1 & "' as NOMBRE, " _
            & "'" & TAXONOMIA1 & "' as TAXONOMIA, '" & AGRUPAMIENTO1 & "' as AGRUPAMIENTO, '" & TIPO_PERSONA1 & "' as TIPO_PERSONA " _
            & "into [" & VF_TP & "];"
            Bandera_Inicio = 1
        ElseIf Bandera = 0 Then
            dbs.Execute "insert into [" & VF_TP & "]  " _
            & "(BANCO, NOMBRE, TAXONOMIA, AGRUPAMIENTO, TIPO_PERSONA) values " _
            & "('" & Banco1 & "', '" & NOMBRE1 & "','" & TAXONOMIA1 & "', '" & AGRUPAMIENTO1 & "', '" & TIPO_PERSONA1 & "');"
        End If
    .MoveNext
    Loop
    End With
    rst.Close
    rst2.Close
    dbs.Close
End Function
'Replace(NOMBRE1, "'", "")
Function ExisteTabla(db_foto_revolventes_cohortes_preliminar, tbl_repetidos_tp_base) As Boolean
    On Error Resume Next
    ExisteTabla = IsObject(OpenDatabase(db_foto_revolventes_cohortes_preliminar).TableDefs(tbl_repetidos_tp_base))
 
End Function

