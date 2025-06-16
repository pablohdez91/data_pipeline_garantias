Sub Carga_BDs_DWH()
    Dim dbs As DAO.database
    Mes = "202411"
    VariableBanco = ""
    Anio = Left(Mes, 4)
    Mes1 = IIf(Mid(Mes, 5, 1) = 0, Right(Mes, 1), Right(Mes, 2))
    Mes2 = Mes_Letra(Mes1) & Mid(Mes, 3, 2)
    
   'Base BD donde estan los parametros de Exportacion
    wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_dwh = wd_processed & "DWH\"
    wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
    wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
    wd_processed_fotos = wd_processed & "Fotos\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"

    'Inputs
    db_catalogos = wd_external & "Catálogos_" & Mes2 & ".accdb"
    db_base = wd_staging & "Base.accdb"
    db_base_vacia = wd_staging & "Base_Vacia.accdb"
    db_cohortes_1 = wd_raw & "Cohortes_" & Mes & "_1.mdb"

    'Outputs
    db_dwh = wd_processed_fotos & "BD_DWH_" & Mes & ".accdb"
    db_dwh_temp = wd_processed_fotos & "BD_DWH_" & Mes & "_Temp.accdb"

    tbl_tipo_credito = "TIPO_CREDITO"
     
    'Se crea la carpeta de Bases del mes procesado
    'Revisa que existe el archivo de no ser así lo crea
    If ExisteRuta(db_dwh) = False Then
       Copia = Application.CompactRepair(db_base, wd_processed_fotos & "Copia de Seguridad.accdb")
       Name wd_processed_fotos & "Copia de Seguridad.accdb" As db_dwh
    End If
    
    'ACOV 201810 cambio en la consulta porque era muy pesada
    'Primero la uno toda en los accesss y despues ya la jalo
    Set dbs = OpenDatabase(db_cohortes_1)
        
        For i = 2 To 5
            dbs.Execute "INSERT INTO DATOS SELECT * FROM DATOS " _
                & " IN '" & wd_raw & "Cohortes_" & Mes & "_" & i & ".mdb' "
        Next i
    dbs.Close
    
    'ACOV 201806 cambio a Tablea
    'Antes se separa en NR Y R la base bajada de DWH en el mismo access dnd se procesará
    Call Vincula_Tabla(db_cohortes_1, db_dwh, "DATOS", "DATOS")
    Call Vincula_Tabla(db_catalogos, db_dwh, tbl_tipo_credito, tbl_tipo_credito)  'Vinculo el tipo de crédito
    Set dbs = OpenDatabase(db_dwh)
        
        dbs.Execute (" SELECT A.BUCKET, A.DESC_INDICADOR AS Producto, A.ESTADO_ID, A.ESTRATO_ID, A.FECHA_APERTURA as [Fecha de Apertura], A.FECHA_PRIMER_INCUMPLIMIENTO, " _
        & " A.FECHA_REGISTRO_ALTA as [Fecha Registro Alta], A.INDICADOR_ID AS [Producto ID], A.INTERMEDIARIO_ID, A.NOMBRE_EMPRESA AS [Empresa / Acreditado (Descripción)], " _
        & " A.NUMERO_CREDITO, A.PLAZO, A.PLAZO_DIAS AS [Plazo Días], A.PORCENTAJE_COMISION_GARANTIA AS [Porcentaje de Comisión Garantia], A.PORCENTAJE_GARANTIZADO AS [Porcentaje Garantizado], " _
        & " A.PROGRAMA_ID, A.PROGRAMA_ORIGINAL, A.RAZON_SOCIAL AS [Razón Social (Intermediario)], A.RFC_EMPRESA AS [RFC Empresa / Acreditado], A.SECTOR_ID, A.TASA_ID, " _
        & " A.TIPO_CREDITO_ID, A.TIPO_GARANTIA_ID, A.VALOR_TASA_INTERES, A.[MONTO_CREDITO_MN (SUMA)] AS [Monto _Credito_Mn], A.CONREC_CLAVE, A.Describe_Desrec, B.NR_R INTO DATOS_VF  " _
        & " FROM DATOS AS A LEFT JOIN " & tbl_tipo_credito & " AS B " _
        & " ON B.Tipo_Credito_ID = A.TIPO_CREDITO_ID; ")
        dbs.Execute " SELECT * INTO BD_DWH_NR_" & Mes & " FROM DATOS_VF WHERE NR_R = 'NR'; "
        dbs.Execute " SELECT * INTO BD_DWH_R_" & Mes & " FROM DATOS_VF WHERE NR_R = 'R'; "
        dbs.Execute " ALTER TABLE BD_DWH_NR_" & Mes & " DROP COLUMN NR_R "
        dbs.Execute " ALTER TABLE BD_DWH_R_" & Mes & " DROP COLUMN NR_R "
        dbs.Execute " DROP TABLE DATOS_VF "
        
    dbs.Close
    
    'Porque sino la base se hace muy pesada
    DAO.DBEngine.CompactDatabase db_dwh, db_dwh_temp
    Kill db_dwh
    Name db_dwh_temp As db_dwh

    'ACOV 201805 Importa el .mdb de Tableau a Access y renombra las columnas a las originales, además corrige los campos
    'Vincula Base
    
For j = 1 To 2
    If j = 1 Then
            Var_NR_R = "NR"
            tbl_bd_dwh_nr_r = "BD_DWH_" & Var_NR_R & "_" & Mes
    Else
            Var_NR_R = "R"
            tbl_bd_dwh_nr_r = "BD_DWH_" & Var_NR_R & "_" & Mes
    End If
        
    
    'ACOV 201805 Corrige los campos con carácteres extraños en el nombre
    If j = 2 Then
    Set dbs = OpenDatabase(db_dwh)
        dbs.Execute ("UPDATE BD_DWH_R_" & Mes & "  SET [Empresa / Acreditado (Descripción)] = IIF( [Empresa / Acreditado (Descripción)] = 'HUGO GUDIO RODRIGUEZ','HUGO GUDINO RODRIGUEZ',[Empresa / Acreditado (Descripción)])")
        dbs.Execute ("UPDATE BD_DWH_R_" & Mes & " SET [Empresa / Acreditado (Descripción)] = IIF( [Empresa / Acreditado (Descripción)] = 'A Y A DISEO SA DE CV','A Y A DISENO SA DE CVZ',[Empresa / Acreditado (Descripción)])")
        dbs.Execute ("UPDATE BD_DWH_R_" & Mes & " SET [Empresa / Acreditado (Descripción)] = IIF( [Empresa / Acreditado (Descripción)] = 'MANOS DEL DIVINO NIO SA DE CV','MANOS DEL DIVINO NINO SA DE CV',[Empresa / Acreditado (Descripción)])")
    dbs.Close
    End If
    
    Set dbs = OpenDatabase(db_dwh)
    
    'Se crea  FVTO_RiesgosD     ([Fecha de Apertura]+round((Plazo/12) *365,0)+(IIF([Plazo Días] is null, 0,[Plazo Días])))
    Call Inserta_Columna(dbs, tbl_bd_dwh_nr_r, "FVTO_Riesgosd", "Date")
    Call Corrige_Campos(dbs, tbl_bd_dwh_nr_r, "FVTO_Riesgosd", "([Fecha de Apertura]+IIF(CDbl(365*[PLAZO]/12)-CInt(365*[PLAZO]/12)>=0.5,CInt(365*[PLAZO]/12)+1,CInt(365*[PLAZO]/12))+(IIF([Plazo Días] is null, 0,[Plazo Días])))", "")
    'Se crea  TPRO_CLAVE
    'Modificar el calculo debido a Emergente
    Call Inserta_Columna(dbs, tbl_bd_dwh_nr_r, "TPRO_CLAVE", "Double")
    Call Corrige_Campos(dbs, tbl_bd_dwh_nr_r, "TPRO_CLAVE", "IIf(A.Programa_Id>=32000 And A.Programa_Id<=32100, A.Programa_Id, IIf(A.Programa_Id=3976 And A.Programa_Original=31415,A.Programa_Id,IIF(A.Programa_Original = 33842 AND A.Programa_Id = 33366, A.Programa_Id, IIF(A.Programa_Original = 3200 AND A.Programa_Id IN (3536, 3537, 3539, 3542,3544, 3545, 3546,3547,3548,3549,3550, 3553, 3555, 3558,3559, 3560, 3564,3566),A.Programa_Id, IIf(A.Programa_Original = 3999,A.Programa_Id,A.Programa_Original))))) ", "")
    'Se crea  'TIPO_PERSONA
    Call Inserta_Columna(dbs, tbl_bd_dwh_nr_r, "TIPO_PERSONA", "Text(2)")
    Call Corrige_Campos(dbs, tbl_bd_dwh_nr_r, "TIPO_PERSONA", "(IIF((Mid([RFC Empresa / Acreditado],4,1)='-'),'M','F'))", "")
    'Se crea  'Nombre_v1
    Call Inserta_Columna(dbs, tbl_bd_dwh_nr_r, "Nombre_v1", "Text(100)")
    Coma = Chr(39)
    
    Call Corrige_Campos(dbs, tbl_bd_dwh_nr_r, "Nombre_v1", "Replace(A.[Empresa / Acreditado (Descripción)],Chr(39),'')", "")
    
Next

End Sub



