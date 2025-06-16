Handler:
    If Err.Number = 3049 Then

        ' 1. Crea copia de la base de datos
        Dim ruta_original As String
        Dim ruta_copia As String
        Dim db_original As String
        
        db_original = dbs.Name
        ruta_original = Left(db_original, Len(db_original) - 6)
        ruta_copia = ruta_original & " - Copia" & ".accdb"

        dbs.Close
        If existe_ruta(ruta_copia) = True Then
            ruta_copia = ruta_original & " - Copia (2)" & ".accdb"
        End If
        DAO.DBEngine.CompactDatabase db_original, ruta_copia
        
        
        ' 2. Elimina Tablas de la BD Original
        Dim tdf As DAO.TableDef
        Dim i As Integer
        Dim tablas_eliminar As Collection
        Set tablas_eliminar = New Collection

        ' Iterar sobre todas las tablas en la colección TableDefs
        Set dbs = OpenDatabase(db_original)
        For Each tdf In dbs.TableDefs
            ' Si la tabla no es una tabla del sistema ni una tabla vinculada (Foreign), se añade a la colección
            If (Left(tdf.Name, 4) <> "MSys") And (tdf.Attributes And dbAttachedTable) = 0 Then
                tablas_eliminar.Add tdf.Name
            End If
        Next tdf
        
        ' Eliminar las tablas locales recogidas
        For i = 1 To tablas_eliminar.Count
            dbs.TableDefs.Delete tablas_eliminar(i)
        Next i
    
    
        ' 3. Vincula tablas eliminadas
        Dim tabla_origen As String
        Dim tabla_destino As String
        
        For Each tabla In tablas_eliminar
            ' Crear un nuevo TableDef y configurar sus propiedades de vínculo
            Set tdf = dbs.CreateTableDef(tabla)
            tdf.Connect = ";DATABASE=" & ruta_copia
            tdf.SourceTableName = tabla
            ' Añadir la tabla vinculada a la base de datos de destino
            dbs.TableDefs.Append tdf
        Next tabla
        
        ' Cerrar la conexión a la base de datos
        dbs.Close
        Set dbs = Nothing

        
        ' 4. Compactar y Reparar
        ' Define la ruta de la base de datos
        Call compacta_repara(db_original)

        Set dbs = OpenDatabase(db_original)
        Resume

    Else
        Dim SaveError As Long
        SaveError = Err.Number
        On Error GoTo 0
        Error (SaveError)
        Resume
    End If