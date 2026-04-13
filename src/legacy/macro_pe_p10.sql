-- pegar_catalogos_valida_saldo
"SELECT I.*, D.Agrupamiento " _
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


-- corrige campos
"update [" & TablaMod & "] " _
    & "set [" & Columna & "] = " & Condicion & " " _
    & "" & Filtro & " ;"


-- Valida saldo
"SELECT A.[descripcion portafolio], A.inter_clave, A.Intermediario_Id, A.[Razón Social], A.tpro_clave_original, A.tpro_clave, A.Progama_Original, A.Programa_Id, Sum(A.[SALDO GARANTIZADO]) AS SumaDeSaldo_contingente_mn " _
          & "INTO [Consulta_valida_saldos_" & Mes & "] " _
          & "FROM [valida_saldos_" & Mes & "] A  " _
          & " GROUP BY A.[descripcion portafolio], A.inter_clave, A.Intermediario_Id, A.[Razón Social], A.tpro_clave_original, A.tpro_clave, A.Progama_Original, A.Programa_Id; "

