 Sub Bases_NR()
    Dim Orden As Long
    Dim fecha_base As Date
    Actual = "202412"
    
    Mes1 = IIf(Mid(Actual, 5, 1) = 0, Right(Actual, 1), Right(Actual, 2))
    Mes_LNTG = Mes_Letra_LNTG(Mes1)
    
    'Mes_LNTG = "B_Febrero"
    fecha_base = DateSerial(Year(Format$(Right(Actual, 2) & "/ 01 /" & Left(Actual, 4), "mm/dd/yyyy")), Month(Format$(Right(Actual, 2) & "/ 01 /" & Left(Actual, 4), "mm/dd/yyyy")) + 1, 0)

    wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_dwh = wd_processed & "DWH\"
    wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
    wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
    wd_processed_fotos = wd_processed & "Fotos\"
    wd_processed_fotos_cierre = wd_processed_fotos & Actual & "\"
    wd_processed_perdida_esperada = wd_processed & "Perdida Esperada\"
    wd_processed_perdida_esperada_bases = wd_processed_perdida_esperada & "Bases\"
    wd_raw = wd & "data\raw\"
    wd_staging = wd & "data\staging\"
    wd_validations = wd & "data\validations\"
    
    tbl_vf_foto_nr = "VF_Foto_NR_" & Actual  '"Foto_201011_VF"
    db_bases_simples = wd_processed_fotos_cierre & "\BasesSimples " & Actual & ".accdb"
    db_foto_simples_cohortes_vf = wd_processed_fotos_cierre & "FotoSimplesCohortes_" & Actual & "_VF.accdb"


    Set Engine = New DBEngine
    Set dbs_root = Engine.CreateDatabase(db_bases_simples, dbLangGeneral)
    dbs_root.Close
    
    Base_Simple_x_Taxo db_bases_simples, db_foto_simples_cohortes_vf, "where (A.[TAXONOMIA]='GARANTIA MICROCREDITO')", tbl_vf_foto_nr, "Microcredito", fecha_base
    Base_Simple_x_Taxo db_bases_simples, db_foto_simples_cohortes_vf, "where (A.[TAXONOMIA]<>'GARANTIA EMPRESARIAL' and A.[TAXONOMIA]<>'GARANTIA MICROCREDITO')", tbl_vf_foto_nr, "Resto", fecha_base
    Base_Simple_x_Taxo db_bases_simples, db_foto_simples_cohortes_vf, "where (A.[TAXONOMIA]='GARANTIA EMPRESARIAL')", tbl_vf_foto_nr, "Empresarial", fecha_base
    
    Agrega_Campos_x_Taxo_NR "Empresarial", db_bases_simples, fecha_base
    Agrega_Campos_x_Taxo_NR "Microcredito", db_bases_simples, fecha_base
    Agrega_Campos_x_Taxo_NR "Resto", db_bases_simples, fecha_base
    
    'Se crea la carpeta de Bases del mes procesado
    If ExisteRuta(Left(wd_processed_perdida_esperada_bases, Len(wd_processed_perdida_esperada_bases) - 6)) = False Then
        MkDir (Left(wd_processed_perdida_esperada_bases, Len(wd_processed_perdida_esperada_bases) - 6))
    End If
    If ExisteRuta(wd_processed_perdida_esperada_bases) = False Then
        MkDir (wd_processed_perdida_esperada_bases)
    End If
    
    Exporta_Access_a_Excel wd_processed_perdida_esperada_bases & "Base_" & "Empresarial" & "_NR_" & Actual & ".xlsx", db_bases_simples, "Empresarial"
    Exporta_Access_a_Excel wd_processed_perdida_esperada_bases & "Base_" & "Resto" & "_NR_" & Actual & ".xlsx", db_bases_simples, "Resto"
    Exporta_Access_a_Excel wd_processed_perdida_esperada_bases & "Base_" & "Microcredito" & "_NR_" & Actual & ".xlsx", db_bases_simples, "Microcredito"
    
End Sub


Function Base_Simple_x_Taxo(db_bases_simples, db_foto_simples_cohortes_vf, Filtro, Foto, TaxoBase, fecha_base As Date)
Dim dbs As DAO.database
Set dbs = OpenDatabase(db_bases_simples)
dbs.Execute "SELECT " _
        & "A.CLAVE_CREDITO, A.FECHA_VALOR1, A.TIPO_PERSONA, A.NOMBRE, A.RFC, A.FECHA_REGISTRO_GARANTIA, A.[MGI (MDP)], " _
        & "A.PLAZO, A.PLAZO_DIAS, A.FVTO, A.BANCO, A.FECHA_PAGO, A.INCUMPLIDO as PAGADAS, A.[MPAGADO (MDP)], A.FECHA_REGISTRO1,  " _
        & "A.[MONTO CREDITO (MDP)], A.FECHA_VALOR, A.INTER_CLAVE, A.TPRO_CLAVE, A.NR_R, A.CSG, A.[SALDO (MDP)], " _
        & "IIF(A.FECHA_PRIMER_INCUM IS NULL, cdate(Format('30/12/1899','dd/mm/yyyy')),A.FECHA_PRIMER_INCUM) as FECHA_PRIMER_INCUM, " _
        & "A.CLAVE_TAXO, A.TAXONOMIA, A.MM_UDIS, IIF(A.CLAVE_CREDITO IS NULL,0,1) as NUM_GAR, A.INCUMPLIDO, A.ESQUEMA, " _
        & "A.[MONTOTOTAL (MDP)], A.[RECUPERADOS (MDP)], A.[RESCATADOS (MDP)], A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, " _
        & "A.PORCENTAJE_GARANTIZADO, A.PLAZO_BUCKET, A.Programa_Original, A.Programa_Id " _
        & "INTO [" & TaxoBase & "] " _
        & "FROM [" & Foto & "] as A " _
        & "IN '" & db_foto_simples_cohortes_vf & "'" _
        & "" & Filtro & ";"
    dbs.Close
    ', PlazoBucket(A.PLAZO, A.PLAZO_DIAS) AS PLAZO_BUCKET
End Function
'Campos que se quitaron...de Simples '''PORCENTAJE_COMISION PLAZO 'PROGRAMA 'CLASIFICACION 'MAS_6MESES_PAGADA 'A2000_A_ACT, MI_5B_RESTO
'Se dejo de Utilizar a partir de Julio 2011 y se utilizo el Plazo tal cual.
'''Function PlazoBucket(Plazo, Plazo_Dias) As Integer
'''    If Plazo_Dias = Null Or Plazo_Dias = 0 Or Plazo_Dias = "" Then
'''        Plazo_Dias1 = 0
'''    Else
'''        Plazo_Dias1 = Plazo_Dias
'''    End If
'''    v_calc = (((Plazo / 12) * 372) + Plazo_Dias1) / 372
'''    If v_calc <= 1 Then
'''        PlazoBucket = 1
'''    ElseIf v_calc > 1 And v_calc <= 2 Then
'''        PlazoBucket = 2
'''    ElseIf v_calc > 2 And v_calc <= 3 Then
'''        PlazoBucket = 3
'''    ElseIf v_calc > 3 Then
'''       PlazoBucket = 4
'''    End If
'''End Function
Function Agrega_Campos_x_Taxo_NR(TablaTaxo, db_bases_simples, fecha_base)
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MGI_VIVOS]", "Double", "(IIF(FVTO+180>#" & fecha_base & "#,1,0)*[MGI (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MGI_MALOS_VIVOS]", "Double", "(IIF(FVTO+180>#" & fecha_base & "#,1,0)*[MGI (MDP)]*INCUMPLIDO)", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MPAGADO_VIVOS]", "Double", "(IIF(FVTO+180>#" & fecha_base & "#,1,0)*[MPAGADO (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MRECUP_VIVOS]", "Double", "(IIF(FVTO+180>#" & fecha_base & "#,1,0)*[RECUPERADOS (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MGI_CAD]", "Double", "(IIF(FVTO+180<=#" & fecha_base & "#,1,0)*[MGI (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MGI_MALOS_CAD]", "Double", "(IIF(FVTO+180<=#" & fecha_base & "#,1,0)*[MGI (MDP)]*INCUMPLIDO)", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MPAGADO_CAD]", "Double", "(IIF(FVTO+180<=#" & fecha_base & "#,1,0)*[MPAGADO (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MRECUP_CAD]", "Double", "(IIF(FVTO+180<=#" & fecha_base & "#,1,0)*[RECUPERADOS (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[SALDO_VIVOS]", "Double", "(IIF(FVTO+180>#" & fecha_base & "#,1,0)*[SALDO (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[SALDO_CADUCOS]", "Double", "(IIF(FVTO+180<=#" & fecha_base & "#,1,0)*[SALDO (MDP)])", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MGI_VIVOS^2]", "Double", "[MGI_VIVOS]*[MGI_VIVOS]", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[#VIVAS]", "Double", "IIF(FVTO+180>#" & fecha_base & "#,1,0)", 0 'ES LA MISMA QUE IND_VIVAS = #vivas
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MGI_INCMPL]", "Double", "[MGI (MDP)]*INCUMPLIDO", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[Count]", "Double", 1, 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[FECHA_PAGO1]", "Double", "IIF(FECHA_PAGO=0,NULL, cdate(Format(dateserial(Year(FECHA_PAGO),Month(FECHA_PAGO),'01'),'dd/mm/yyyy')))", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[AñoOtor]", "Double", "YEAR(FECHA_VALOR1)", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[PTRANSCURRIDO]", "Double", "(IIF(1>((#" & fecha_base & "# - FECHA_VALOR)/(FVTO-FECHA_VALOR+180)),((#" & fecha_base & "# - FECHA_VALOR)/(FVTO-FECHA_VALOR+180)),1))", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[PTRANS_PON]", "Double", "[PTRANSCURRIDO]*[MGI (MDP)]", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[Semestre]", "Double", "IIF(MONTH(FECHA_VALOR)<7,cdate(Format(dateserial(YEAR(FECHA_VALOR),'01','01'),'dd/mm/yyyy')),cdate(Format(dateserial(YEAR(FECHA_VALOR),'02','01'),'dd/mm/yyyy')))", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[MESES_REM_POND]", "Double", "(1-PTRANSCURRIDO)*((YEAR(FVTO)-YEAR(FECHA_VALOR1))*12+(MONTH(FVTO)-MONTH(FECHA_VALOR1))+6)*[SALDO (MDP)]", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[Con_Saldo]", "Double", "IIF([SALDO (MDP)]>0,1,0)", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[RESTANTE_MESES]", "Double", "IIF(0<(IIF([Con_Saldo]=1,(YEAR(FVTO)-YEAR(#" & fecha_base & "#))*12+(MONTH(FVTO)-MONTH(#" & fecha_base & "#)),0)),(IIF([Con_Saldo]=1,(YEAR(FVTO)-YEAR(#" & fecha_base & "#))*12+(MONTH(FVTO)-MONTH(#" & fecha_base & "#)),0)),0)", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[RESTANTE_POND]", "Double", "RESTANTE_MESES*[SALDO (MDP)]", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[SALDO^2]", "Double", " [SALDO (MDP)] * [SALDO (MDP)]", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[RESTANTE_DIAS]", "Double", "IIF([SALDO (MDP)]>0,(FVTO-#" & fecha_base & "#),0)", 0
    Campos_Extras_Bases_NR TablaTaxo, db_bases_simples, "[RESTANTE_DIAS_POND]", "Double", "RESTANTE_DIAS*[SALDO (MDP)]", 1
End Function
Function Campos_Extras_Bases_NR(TaxoBase, db_foto_simples_cohortes_vf, Columna, TP_Columna, Condicion, Bandera)
    Dim dbs As database
    Set dbs = OpenDatabase(db_foto_simples_cohortes_vf)
    dbs.Execute "alter table " & TaxoBase & " " _
                & "add column " & Columna & " " & TP_Columna & " ; "
                dbs.Execute "update " & TaxoBase & " " _
                & "set " & Columna & " = " & Condicion & " ;"
    
    If Bandera = 1 Then
        dbs.Close
    End If
End Function
Function Campos_Extras_Bases_NR2(db_bases_simples, db_foto_simples_cohortes_vf, Linea, Filtro, Foto, TaxoBase, fecha_base)
    Dim dbs As database
    Set dbs = OpenDatabase(db_bases_simples)
    dbs.Execute "SELECT " & Linea & ", " _
        & "IIF(A.[SALDO (MDP)]>0,1,0) as [No_Acreditados_Saldo>0], A.[SALDO (MDP)]*A.[SALDO (MDP)] as Saldo^2, 1 as Count, " _
        & "IIF(A.[SALDO (MDP)]>0,IIF(A.ANTIG_CLIENTE_MESES<=12,1,IIF(A.ANTIG_CLIENTE_MESES<=24,2,IIF(A.ANTIG_CLIENTE_MESES<=36,3,4))),0) AS ANTIG_CLIENTE_AÑOS, " _
        & "A.RESTANTE_MESES*A.[SALDO (MDP)] as RESTANTE_POND, IIF(A.FVTO+180> #" & fecha_base & "#,1,0) as VIGENTES, " _
        & "IIF(A.REMANENTE_MESES<=12,1,IIF(A.REMANENTE_MESES<=24,2,IIF(A.REMANENTE_MESES<=36,3,IIF(A.REMANENTE_MESES<=48,4,5)))) as REMANENTE_AÑOS" _
        & "IIF(A.REMANENTE_MESES+180<=12,1,IIF(A.REMANENTE_MESES+180<=24,2,IIF(A.REMANENTE_MESES+180<=36,3,IIF(A.REMANENTE_MESES+180<=48,4,5)))) as REMANENTE_AÑOS+180" _
        & "A.ANTIG_CLIENTE_MESES * A.[SALDO (MDP)] AS Antig_Cliente_Meses_Pond" _
        & "INTO [" & TaxoBase & "] " _
        & "FROM [" & Foto & "] as A " _
        & "IN '" & db_foto_simples_cohortes_vf & "';"
    dbs.Close
End Function
Function Pega_Nombres_S(db_bases_simples, db_foto_simples_cohortes_vf, Orden, Anio, Mes)
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
          & "INTO [" & Foto & "_F] " _
          & "FROM [" & Foto & "] AS A " _
          & "IN '" & db_foto_simples_cohortes_vf & "';"
    dbs.Close
End Function


