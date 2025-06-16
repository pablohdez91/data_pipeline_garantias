Sub Bases_R()
    Dim Orden As Long
    'Dim fecha_base As Date
    Actual = "202412"
    
    Mes1 = IIf(Mid(Actual, 5, 1) = 0, Right(Actual, 1), Right(Actual, 2))
    Mes_LNTG = Mes_Letra_LNTG(Mes1)
    
    'Mes = Month(Format$(Right(Actual, 2) & "/ 01 /" & Left(Actual, 4), "mm/dd/yyyy")) + 1
    'Anio = Year(Format$(Right(Actual, 2) & "/ 01 /" & Left(Actual, 4), "mm/dd/yyyy"))
    'fecha_base = DateSerial(Anio, 13, 0)
    'fecha_base = DateSerial(Year(Format$(Right(Actual, 2) & "/ 01 /" & Left(Actual, 4), "mm/dd/yyyy")), Month(Format$(Right(Actual, 2) & "/ 01 /" & Left(Actual, 4), "mm/dd/yyyy")) + 1, 0)
    fecha_base = "31/12/2024"
    
    

    wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_dwh = wd_processed & "DWH\"
    wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
    wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
    wd_processed_fotos = wd_processed & "Fotos\"
    wd_processed_fotos_cierre = wd_processed_fotos & Actual & "\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"


    db_foto_revolventes_cohortes_preliminar = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & Actual & "_Preeliminar.accdb"
    db_bases_revolventes = wd_processed_fotos_cierre & "BasesRevolventes " & Actual & ".accdb"


    tbl_vf_foto_r = "VF_Foto_R_" & Actual  '"Foto_201011_VF"
    Linea = "A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.NOMBRE, A.TAXONOMIA, A.CLAVE_TAXO, A.NR_R, A.INTER_CLAVE, A.CSG "
    

    Set Engine = New DBEngine
    Set dbs_root = Engine.CreateDatabase(db_bases_revolventes, dbLangGeneral)
    dbs_root.Close
    
    Agrupa_x_Taxo db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, "where (A.[TAXONOMIA]='GARANTIA MICROCREDITO')", tbl_vf_foto_r, "Microcredito"
    Agrupa_x_Taxo db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, "where (A.[TAXONOMIA]<>'GARANTIA EMPRESARIAL' and A.[TAXONOMIA]<>'EMPRESA MEDIANA' and A.[TAXONOMIA]<>'GARANTIA MICROCREDITO')", tbl_vf_foto_r, "Resto"
    Agrupa_x_Taxo db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, "where (A.[TAXONOMIA]='GARANTIA EMPRESARIAL' or A.[TAXONOMIA]='EMPRESA MEDIANA')", tbl_vf_foto_r, "Empresarial"
    Genera_x_Taxo db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, "where (A.[TAXONOMIA]='GARANTIA EMPRESARIAL')", tbl_vf_foto_r, "Empresarial_x_Disposicion"
    'Empresarial_x_Disposicion
    Agrega_Campos_x_Taxo "Empresarial_x_Disposicion", db_bases_revolventes, fecha_base
    Agrega_Campos_x_Taxo "Empresarial", db_bases_revolventes, fecha_base
    Agrega_Campos_x_Taxo "Microcredito", db_bases_revolventes, fecha_base
    Agrega_Campos_x_Taxo "Resto", db_bases_revolventes, fecha_base
End Sub


Function Agrupa_x_Taxo(db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, Filtro, tbl_vf_foto_r, TaxoBase)
    Dim dbs As database
    Set dbs = OpenDatabase(db_bases_revolventes)
    dbs.Execute "SELECT " & Linea & ", " _
        & "MIN(A.FECHA_VALOR1) as FECHA_VALOR1, MIN(A.FECHA_REGISTRO_GARANTIA) as FECHA_REGISTRO_GARANTIA, " _
        & "SUM(A.[MGI (MDP)]) as [MGI (MDP)], AVG(A.PLAZO) as PLAZO, AVG(A.PLAZO_DIAS) as PLAZO_DIAS, MAX(A.FVTO) as FVTO, " _
        & "Min(A.FECHA_PAGO) as FECHA_PAGO, SUM(A.INCUMPLIDO) as PAGADAS, SUM(A.[MPAGADO (MDP)]) as [MPAGADO (MDP)], " _
        & "SUM(A.[MONTO CREDITO (MDP)]) as [MONTO CREDITO (MDP)], MIN(A.FECHA_VALOR) as FECHA_VALOR, " _
        & "SUM(A.[SALDO (MDP)]) as [SALDO (MDP)], MIN(A.FECHA_REGISTRO1) as FECHA_REGISTRO1, " _
        & "IIF(Min(A.FECHA_PRIMER_INCUM) IS NULL, cdate(Format('30/12/1899','dd/mm/yyyy')), Min(A.FECHA_PRIMER_INCUM)) as FECHA_PRIMER_INCUM, " _
        & "MAX(A.MM_UDIS) as MM_UDIS, COUNT(A.NUM_GAR) as NUM_GAR, MAX(A.INCUMPLIDO) as INCUMPLIDO, First(A.ESQUEMA) as ESQUEMA, " _
        & "SUM(A.[MONTOTOTAL (MDP)]) as [MONTOTOTAL (MDP)], SUM(A.[RECUPERADOS (MDP)]) as [TOT RECUP (MDP)], SUM(A.[RESCATADOS (MDP)]) as [TOT RESCAT (MDP)] " _
        & "INTO [" & TaxoBase & "] " _
        & "FROM [" & tbl_vf_foto_r & "] as A " _
        & "IN '" & db_foto_revolventes_cohortes_preliminar & "'" _
        & "" & Filtro & "" _
        & "GROUP BY  " & Linea & " " _
        & "ORDER BY " & Linea & "; "
    dbs.Close
End Function
Function Genera_x_Taxo(db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, Filtro, tbl_vf_foto_r, TaxoBase)
    Dim dbs As database
    Set dbs = OpenDatabase(db_bases_revolventes)
    dbs.Execute "SELECT " & Linea & ", " _
        & "A.FECHA_VALOR1, A.FECHA_REGISTRO_GARANTIA, " _
        & "A.[MGI (MDP)], A.PLAZO, A.PLAZO_DIAS, A.FVTO, " _
        & "A.FECHA_PAGO, A.INCUMPLIDO as PAGADAS, A.[MPAGADO (MDP)], " _
        & "A.[MONTO CREDITO (MDP)],A.FECHA_VALOR, " _
        & "A.[SALDO (MDP)], A.FECHA_REGISTRO1, " _
        & "IIF(A.FECHA_PRIMER_INCUM IS NULL, cdate(Format('30/12/1899','dd/mm/yyyy')), A.FECHA_PRIMER_INCUM) as FECHA_PRIMER_INCUM, " _
        & "A.MM_UDIS, A.NUM_GAR, A.INCUMPLIDO, " _
        & "A.[MONTOTOTAL (MDP)], A.[RECUPERADOS (MDP)] as [TOT RECUP (MDP)], A.[RESCATADOS (MDP)] as [TOT RESCAT (MDP)] " _
        & "INTO [" & TaxoBase & "] " _
        & "FROM [" & tbl_vf_foto_r & "] as A " _
        & "IN '" & db_foto_revolventes_cohortes_preliminar & "'" _
        & "" & Filtro & ";"
    dbs.Close
End Function
Function Agrega_Campos_x_Taxo(TablaTaxo, db_bases_revolventes, fecha_base)
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[No_Acreditados_Saldo>0]", "Double", "IIF([SALDO (MDP)]>0,1,0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[Saldo^2]", "Double", "[SALDO (MDP)]*[SALDO (MDP)]", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[Count]", "Double", 1, 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[ANTIG_CLIENTE_MESES]", "Double", "IIF([SALDO (MDP)]>0,(YEAR(#" & fecha_base & "#)-YEAR(FECHA_VALOR1))*12+(MONTH(#" & fecha_base & "#)-MONTH(FECHA_VALOR1)),0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[ANTIG_CLIENTE_AÑOS]", "Double", "IIF([SALDO (MDP)]>0,IIF(ANTIG_CLIENTE_MESES<=12,1,IIF(ANTIG_CLIENTE_MESES<=24,2,IIF(ANTIG_CLIENTE_MESES<=36,3,4))),0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[RESTANTE_MESES]", "Double", "IIF((IIF([SALDO (MDP)]>0,(YEAR(FVTO)-YEAR(#" & fecha_base & "#))*12+(MONTH(FVTO)-MONTH(#" & fecha_base & "#)),0))<0,0,(IIF([SALDO (MDP)]>0,(YEAR(FVTO)-YEAR(#" & fecha_base & "#))*12+(MONTH(FVTO)-MONTH(#" & fecha_base & "#)),0)))", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[RESTANTE_POND]", "Double", "RESTANTE_MESES*[SALDO (MDP)]", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[VIGENTES]", "Double", "IIF(FVTO+180> #" & fecha_base & "#,1,0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[REMANENTE_MESES]", "Double", "IIF([SALDO (MDP)]>0,IIF(FVTO> #" & fecha_base & "#,(YEAR(FVTO)-YEAR(#" & fecha_base & "#))*12+(MONTH(FVTO)-MONTH(#" & fecha_base & "#)),0),0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[REMANENTE_AÑOS]", "Double", "IIF(REMANENTE_MESES<=12,1,IIF(REMANENTE_MESES<=24,2,IIF(REMANENTE_MESES<=36,3,IIF(REMANENTE_MESES<=48,4,5))))", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[REMANENTE_MESES+180]", "Double", "IIF([SALDO (MDP)]>0,IIF(FVTO+180>#" & fecha_base & "#,(YEAR(FVTO)-YEAR(#" & fecha_base & "#))*12+(MONTH(FVTO)-MONTH(#" & fecha_base & "#))+180/30,0),0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[REMANENTE_AÑOS+180]", "Double", "IIF(REMANENTE_MESES+180<=12,1,IIF(REMANENTE_MESES+180<=24,2,IIF(REMANENTE_MESES+180<=36,3,IIF(REMANENTE_MESES+180<=48,4,5))))", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[Antig_Cliente_Meses_Pond]", "Double", "ANTIG_CLIENTE_MESES * [SALDO (MDP)]", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[RESTANTE_DIAS]", "Double", "IIF([SALDO (MDP)]>0,(FVTO-#" & fecha_base & "#),0)", 0
    Campos_Extras_Bases_2 TablaTaxo, db_bases_revolventes, "[RESTANTE_DIAS_POND]", "Double", "RESTANTE_DIAS*[SALDO (MDP)]", 1
End Function
Function Campos_Extras_Bases_2(TaxoBase, db_foto_revolventes_cohortes_preliminar, Columna, TP_Columna, Condicion, Bandera)
    Dim dbs As database
    Set dbs = OpenDatabase(db_foto_revolventes_cohortes_preliminar)
    dbs.Execute "alter table " & TaxoBase & " " _
                & "add column " & Columna & " " & TP_Columna & " ; "
                dbs.Execute "update " & TaxoBase & " " _
                & "set " & Columna & " = " & Condicion & " ;"
    
    
    If Bandera = 1 Then
        dbs.Close
    End If
End Function
Function Campos_Extras_Bases(db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Linea, Filtro, tbl_vf_foto_r, TaxoBase, fecha_base)
    Dim dbs As database
    Set dbs = OpenDatabase(db_bases_revolventes)
    dbs.Execute "SELECT " & Linea & ", " _
        & "IIF(A.[SALDO (MDP)]>0,1,0) as [No_Acreditados_Saldo>0], A.[SALDO (MDP)]*A.[SALDO (MDP)] as Saldo^2, 1 as Count, " _
        & "IIF(A.[SALDO (MDP)]>0,IIF(A.ANTIG_CLIENTE_MESES<=12,1,IIF(A.ANTIG_CLIENTE_MESES<=24,2,IIF(A.ANTIG_CLIENTE_MESES<=36,3,4))),0) AS ANTIG_CLIENTE_AÑOS, " _
        & "A.RESTANTE_MESES*A.[SALDO (MDP)] as RESTANTE_POND, IIF(A.FVTO+180> #" & fecha_base & "#,1,0) as VIGENTES, " _
        & "IIF(A.REMANENTE_MESES<=12,1,IIF(A.REMANENTE_MESES<=24,2,IIF(A.REMANENTE_MESES<=36,3,IIF(A.REMANENTE_MESES<=48,4,5)))) as REMANENTE_AÑOS" _
        & "IIF(A.REMANENTE_MESES+180<=12,1,IIF(A.REMANENTE_MESES+180<=24,2,IIF(A.REMANENTE_MESES+180<=36,3,IIF(A.REMANENTE_MESES+180<=48,4,5)))) as REMANENTE_AÑOS+180" _
        & "A.ANTIG_CLIENTE_MESES * A.[SALDO (MDP)] AS Antig_Cliente_Meses_Pond" _
        & "INTO [" & TaxoBase & "] " _
        & "FROM [" & tbl_vf_foto_r & "] as A " _
        & "IN '" & db_foto_revolventes_cohortes_preliminar & "';"
    dbs.Close
End Function
Function Pega_nombres(db_bases_revolventes, db_foto_revolventes_cohortes_preliminar, Orden, Anio, Mes)
    Dim dbs As database
    Set dbs = OpenDatabase(Base)
    dbs.Execute "SELECT A.*, " _
          & "IIF(A.FECHA_PRIMER_INCUM=0,NULL,cdate(Format(dateserial(Year(A.FECHA_PRIMER_INCUM),Month(A.FECHA_PRIMER_INCUM),'01'),'dd/mm/yyyy'))) AS FECHA_PRIMER_INCUM1, " _
          & "IIF(A.FECHA_PAGO=0,NULL, cdate(Format(dateserial(Year(A.FECHA_PAGO),Month(A.FECHA_PAGO),'01'),'dd/mm/yyyy'))) AS FECHA_PAGO1, " _
          & "IIF(DateSerial(" & Anio & "," & Mes & " + 1, 1)<= A.FECHA_REGISTRO_GARANTIA,0,A.MM_UDIS) as MM_UDIS_A, " _
          & "IIF(DateSerial(" & Anio & "," & Mes & " + 1, 1)<= A.FECHA_REGISTRO_GARANTIA, NULL, IIF(A.FVTO=0, NULL, cdate(Format(dateserial(Year(A.FVTO),Month(A.FVTO),'01'),'dd/mm/yyyy')))) AS FVTO1, " _
          & "IIF(DateSerial(" & Anio & "," & Mes & " + 1, 1)<= A.FECHA_REGISTRO_GARANTIA, NULL, cdate(Format(dateserial(Year(A.FECHA_VALOR1),Month(A.FECHA_VALOR1),day(A.FECHA_VALOR1)),'dd/mm/yyyy'))) AS FECHA_VALOR_A, " _
          & "IIF(DateSerial(" & Anio & "," & Mes & " + 1, 1)<= A.FECHA_REGISTRO_GARANTIA, NULL, cdate(Format(dateserial(Year(A.FECHA_REGISTRO_GARANTIA),Month(A.FECHA_REGISTRO_GARANTIA),day(A.FECHA_REGISTRO_GARANTIA),'dd/mm/yyyy'))) AS FECHA_REGISTRO_GARANTIA_A, " _
          & "IIF(DateSerial(" & Anio & "," & Mes & " + 1, 1)<= A.FECHA_REGISTRO_GARANTIA, NULL, cdate(Format(dateserial(Year(A.FECHA_VALOR1),Month(A.FECHA_VALOR1),'01'),'dd/mm/yyyy'))) AS FECHA_VALOR1_A, " _
          & "IIF(DateSerial(" & Anio & "," & Mes & " + 1, 1)<= A.FECHA_REGISTRO_GARANTIA, NULL, cdate(Format(dateserial(Year(A.FECHA_REGISTRO_GARANTIA),Month(A.FECHA_REGISTRO_GARANTIA),'01'),'dd/mm/yyyy'))) AS FECHA_REGISTRO1_A " _
          & "INTO [" & tbl_vf_foto_r & "_F] " _
          & "FROM [" & tbl_vf_foto_r & "] AS A " _
          & "IN '" & db_foto_revolventes_cohortes_preliminar & "';"
    dbs.Close
End Function
 Function Mes_Letra_LNTG(Mes) As String
    Select Case Mes
    Case 1
        Mes_Letra_LNTG = "A_Enero"
    Case 2
        Mes_Letra_LNTG = "B_Febrero"
    Case 3
        Mes_Letra_LNTG = "C_Marzo"
    Case 4
        Mes_Letra_LNTG = "D_Abril"
    Case 5
        Mes_Letra_LNTG = "E_Mayo"
    Case 6
        Mes_Letra_LNTG = "F_Junio"
    Case 7
        Mes_Letra_LNTG = "G_Julio"
    Case 8
        Mes_Letra_LNTG = "H_Agosto"
    Case 9
        Mes_Letra_LNTG = "I_Septiembre"
    Case 10
        Mes_Letra_LNTG = "J_Octubre"
    Case 11
        Mes_Letra_LNTG = "K_Noviembre"
    Case 12
        Mes_Letra_LNTG = "L_Diciembre"
    End Select
 End Function


update Empresarial_x_Disposicion
set REMANENTE_MESES = "IIF([SALDO (MDP)]>0,IIF(FVTO> #31/12/2024#,(YEAR(FVTO)-YEAR(#31/12/2024#))*12+(MONTH(FVTO)-MONTH(#31/12/2024#)),0),0)"