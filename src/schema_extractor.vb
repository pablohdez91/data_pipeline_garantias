Sub ExportarSchemaINI()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim schemaText As String
    Dim fileName As String
    Dim i As Integer

    ' Nombre de la tabla Access y del archivo CSV correspondiente
    Dim nombreTabla As String
    nombreTabla = "Pagadas_Global_VF_202503"  ' ? cambia esto por el nombre de tu tabla
    fileName = "Pagadas_Global_VF_202503.csv" ' ? y esto por el nombre real del archivo CSV

    Set db = CurrentDb
    Set tdf = db.TableDefs(nombreTabla)

    schemaText = "[" & fileName & "]" & vbCrLf
    schemaText = schemaText & "Format=CSVDelimited" & vbCrLf
    schemaText = schemaText & "ColNameHeader=True" & vbCrLf
    schemaText = schemaText & "CharacterSet=ANSI" & vbCrLf

    i = 1
    For Each fld In tdf.Fields
        schemaText = schemaText & "Col" & i & "=" & fld.Name & " " & AccessToSchemaIniType(fld.Type, fld.Size) & vbCrLf
        i = i + 1
    Next fld

    ' Guardar en archivo schema.ini
    Dim fso As Object, archivo As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ruta = "E:\Users\jhernandezr\DAR\garantias\data_pipeline_garantias\data\processed\DWH\"
    Set archivo = fso.CreateTextFile(ruta & "schema.ini", True) ' ? cambia esta ruta
    archivo.Write schemaText
    archivo.Close

    MsgBox "Schema.ini generado con éxito."
End Sub

Function AccessToSchemaIniType(tipo As Integer, tamano As Integer) As String
    ' Mapea los tipos de Access a los tipos de schema.ini
    Select Case tipo
        Case 1
            AccessToSchemaIniType = "Boolean"
        Case 2
            AccessToSchemaIniType = "Byte"
        Case 3
            AccessToSchemaIniType = "Integer"
        Case 4
            AccessToSchemaIniType = "Long Integer"
        Case 5
            AccessToSchemaIniType = "Currency"
        Case 6
            AccessToSchemaIniType = "Single"
        Case 7
            AccessToSchemaIniType = "Double"
        Case 8
            AccessToSchemaIniType = "Date"
        Case 10
            AccessToSchemaIniType = "Text Width " & tamano
        Case 12
            AccessToSchemaIniType = "Memo"
        Case Else
            AccessToSchemaIniType = "Text"
    End Select
End Function

