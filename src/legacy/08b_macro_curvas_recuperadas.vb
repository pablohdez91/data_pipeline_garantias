Function Cruza_Llave(dbs, T_Final, T_Izquierda, T_Derecha)
odbs.Execute "SELECT A.*, B.LLAVE_FINAL " _
          & "INTO [" & T_Final & "] " _
          & "FROM [" & T_Izquierda & "] as A LEFT JOIN [" & T_Derecha & "] as B " _
          & "ON (A.[TPRO_CLAVE]*1)=B.[Programa Mezcla ID] AND A.Intermediario_Id=B.Intermediario_Id; "
End Function

Function Base_Recup_Detalle(dbs, T_Final, T_Izquierda, T_Derecha)
dbs.Execute "SELECT A.*, " _
    & "IIF(A.Monto_Mn is null,0,A.Monto_Mn)/1000000 AS [CAPRECUP(MDP)], " _
    & "IIF(A.Interes_Mn is null,0,A.Interes_Mn)/1000000 AS [INTRECUP(MDP)], IIF(A.Moratorios_Mn is null,0,A.Moratorios_Mn)/1000000 AS [MORAT(MDP)], " _
    & "IIF(A.Excedente_Mn is null,0,A.Excedente_Mn)/1000000 AS [EXCEDENTE(MDP)], " _
    & "IIF(A.[Gastos_Juicio_Mn] is null,0,A.[Gastos_Juicio_Mn])/1000000 AS [GASTOSJUICIO(MDP)], " _
    & "IIF(A.[Sub_Total_Mn] is null,0,A.[Sub_Total_Mn])/1000000 AS [SUBTOTAL (MDP)], " _
    & "IIF(A.Monto_Total_Mn is null,0,A.Monto_Total_Mn)/1000000 AS [MONTOTOTAL (MDP)], " _
    & "IIF(A.Estatus='E','E','NE') AS REEST, B.Estatus as STATUS_RECUP, " _
    & "IIF(A.Fecha > Fecha_Garantia_Honrada,1,0) AS IND_ENTRA, " _
    & "IIF(B.Estatus ='Recuperada',(IIF(A.Monto_Mn is null,0,A.Monto_Mn)+IIF(A.Interes_Mn is null,0,A.Interes_Mn)+IIF(A.Moratorios_Mn is null,0,A.Moratorios_Mn)+IIF(A.Excedente_Mn is null,0,A.Excedente_Mn)-IIF(A.[Gastos_Juicio_Mn] is null,0,A.[Gastos_Juicio_Mn]))/1000000,0) AS [MRECUP_TOT(MDP)], " _
    & "IIF(B.Estatus ='Rescatada',(IIF(A.Monto_Mn is null,0,A.Monto_Mn)+IIF(A.Interes_Mn is null,0,A.Interes_Mn)+IIF(A.Moratorios_Mn is null,0,A.Moratorios_Mn)+IIF(A.Excedente_Mn is null,0,A.Excedente_Mn)-IIF(A.[Gastos_Juicio_Mn] is null,0,A.[Gastos_Juicio_Mn]))/1000000,0) AS [MRESCAT_TOT(MDP)], " _
    & "(YEAR(A.FECHA)-YEAR(A.Fecha_Garantia_Honrada))*12+(MONTH(A.FECHA)-MONTH(A.Fecha_Garantia_Honrada)) AS DIFERENCIAS_MESES, " _
    & "(YEAR(A.Fecha_Garantia_Honrada)) AS ANIO_PAGO, " _
    & "(YEAR(A.Fecha_Garantia_Honrada))*100+(MONTH(A.Fecha_Garantia_Honrada)) AS ORDEN_PAGO, " _
    & "A.Numero_Credito & A.Intermediario_Id AS CONCATENAR_SALDOS " _
    & "INTO " & T_Final & " " _
    & "FROM " & T_Izquierda & " as A left join " & T_Derecha & " as B " _
    & "ON (A.Estatus=B.[Estatus ID])  ; "
End Function

Function CiclosRescate(dbs, T_Final, T_Inicial)
dbs.Execute "SELECT CONCATENAR_SALDOS, Numero_Credito as Prestamo, [Razón Social (Intermediario)] as INTERMEDIARIO, MRESCAT_TOTAL, " _
    & "IIF(MRESCAT_TOTAL=0,1,0) AS CICLOS_RESCAT " _
    & "INTO [" & T_Final & "] " _
    & " FROM (SELECT Numero_Credito, [Razón Social (Intermediario)], CONCATENAR_SALDOS, SUM([MRESCAT_TOT(MDP)]*IND_ENTRA) AS MRESCAT_TOTAL " _
          & " FROM [" & T_Inicial & "] " _
          & " WHERE (Numero_Credito is not null) and ([Razón Social (Intermediario)] is not null) " _
          & " GROUP BY Numero_Credito, [Razón Social (Intermediario)], CONCATENAR_SALDOS); "
End Function

Function Cruza_CiclosRescate(dbs, T_Final, T_Izquiera, T_Derecha)

dbs.Execute "SELECT A.*, B.CICLOS_RESCAT " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Izquiera & "] as A LEFT JOIN [" & T_Derecha & "] as B " _
    & "ON A.CONCATENAR_SALDOS = B.CONCATENAR_SALDOS; "
End Function

Function AgrupaRecup(dbs, T_Final, T_Inicial, CampoExtra, CampoExtraAgrup)
dbs.Execute "SELECT IND_ENTRA, MM_UDIS, Producto AS TAXONOMIA, NR_R, LLAVE_FINAL, TPRO_CLAVE, " _
    & "[Razón Social (Intermediario)] AS BANCO, CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, Intermediario_Id, CICLOS_RESCAT, " _
    & "DIFERENCIAS_MESES, Tipo_Garantia_Id AS TIPGAR_CLAVE, ANIO_PAGO AS [ANIO(PAGO)], REEST, " _
    & "YEAR(Fecha_Registro) as Anio_REG_RECUP, " & CampoExtra & " ESQUEMA, ORDEN_PAGO, " _
    & "SUM([MRECUP_TOT(MDP)]) AS MRECUP_TOT, SUM([MRESCAT_TOT(MDP)]) AS MRESCAT_TOT " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "] " _
    & "GROUP BY IND_ENTRA, MM_UDIS, PRODUCTO, NR_R, LLAVE_FINAL, TPRO_CLAVE, [Razón Social (Intermediario)], " _
    & "CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, Intermediario_Id, CICLOS_RESCAT, DIFERENCIAS_MESES, Tipo_Garantia_Id, " _
    & "ANIO_PAGO, REEST, YEAR(Fecha_Registro), " & CampoExtraAgrup & " ESQUEMA, ORDEN_PAGO ;"
End Function

Function Agrupa_Recup_VF_Completo(dbs, T_Final, T_Inicial)
dbs.Execute "SELECT *, TAXONOMIA, CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, " _
    & "IIF(ESQUEMA='PP','Pari Passu',IIF(ESQUEMA is null,'Pari Passu','1P')) AS ESQUEMA_VF, " _
    & "IIF([ANIO(PAGO)]<=2006,2006, [ANIO(PAGO)]) AS [ANIO(PAGO)_New], LLAVE_FINAL, TPRO_CLAVE " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "]; "
End Function

Function Agrupa_Recup_VF_Extracto(dbs, T_Final, T_Inicial)
dbs.Execute "SELECT IND_ENTRA, MM_UDIS, TAXONOMIA, NR_R, BANCO, CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, " _
    & "CICLOS_RESCAT, DIFERENCIAS_MESES, TIPGAR_CLAVE, [ANIO(PAGO)], REEST, " _
    & "MRECUP_TOT, MRESCAT_TOT, Anio_REG_RECUP, " _
    & "ESQUEMA_VF AS ESQUEMA, [ANIO(PAGO)_New], Intermediario_Id, TPRO_CLAVE, ORDEN_PAGO " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "]; "
End Function

Function Agrupa_UltimoPago(dbs, T_Final, T_Inicial)
dbs.Execute "SELECT Intermediario_Id, Numero_Credito, MAX([Pago ID]) AS ULT_PAGO  " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "] " _
    & "GROUP BY Intermediario_Id, Numero_Credito; "
End Function

Function Toma_UltimoPago(dbs, T_Final, T_Izquierda, T_Derecha)
dbs.Execute "SELECT A.*, B.ULT_PAGO " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Izquierda & "] as A LEFT JOIN [" & T_Derecha & "] as B " _
    & "ON A.Intermediario_Id=B.Intermediario_Id AND A.Numero_Credito=B.Numero_Credito AND A.[Pago ID]=B.ULT_PAGO " _
    & "WHERE B.ULT_PAGO IS NOT NULL;"
End Function

Function Base_Pagadas_Detalle(dbs, T_Final, T_Inicio)
dbs.Execute "SELECT A.*, " _
    & "IIF(A.Monto_Pagado_Mn is null,0,A.Monto_Pagado_Mn)/1000000 AS [MPAGADO (MDP)], " _
    & "(YEAR(A.Fecha_Garantia_Honrada)) AS ANIO_PAGO, " _
    & "(YEAR(A.Fecha_Garantia_Honrada))*100+MONTH(A.Fecha_Garantia_Honrada) AS ORDEN_PAGO , " _
    & "A.[Numero_Credito] & A.Intermediario_Id AS CONCATENAR_SALDOS " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicio & "] as A ;"
End Function

Function Cruza_CiclosRescate(dbs, T_Final, T_Izquiera, T_Derecha)
dbs.Execute "SELECT A.*, B.CICLOS_RESCAT " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Izquiera & "] as A LEFT JOIN [" & T_Derecha & "] as B " _
    & "ON A.CONCATENAR_SALDOS = B.CONCATENAR_SALDOS; "
End Function

dbs.Execute "SELECT MM_UDIS, PRODUCTO AS TAXONOMIA, NR_R, LLAVE_FINAL, TPRO_CLAVE, " _
    & "[Razón Social (Intermediario)] AS BANCO, CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, Intermediario_Id, CICLOS_RESCAT, " _
    & "Tipo_Garantia_Id AS TIPGAR_CLAVE, ANIO_PAGO AS [ANIO(PAGO)], ESQUEMA, ORDEN_PAGO, " _
    & "SUM([MPAGADO (MDP)]) AS MPAGADO " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "] " _
    & "GROUP BY MM_UDIS, PRODUCTO, NR_R, LLAVE_FINAL, TPRO_CLAVE, [Razón Social (Intermediario)], " _
    & "CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, Intermediario_Id, CICLOS_RESCAT, Tipo_Garantia_Id, ANIO_PAGO, ESQUEMA, ORDEN_PAGO;"
End Function

Function Agrupa_Pagados_VF_Completo(dbs, T_Final, T_Inicial)
dbs.Execute "SELECT *, TAXONOMIA, CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, " _
    & "IIF(CICLOS_RESCAT is null,1,IIF(CICLOS_RESCAT=0,0,1)) AS  CICLOS_RESCAT_VF, " _
    & "[ANIO(PAGO)], LLAVE_FINAL, TPRO_CLAVE,  " _
    & "IIF(ESQUEMA='PP','Pari Passu',IIF(ESQUEMA is null,'Pari Passu','1P')) AS ESQUEMA_VF, " _
    & "IIF([ANIO(PAGO)]<=2006,2006, [ANIO(PAGO)]) AS [ANIO(PAGO)_New] " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "]; "
End Function

Function Agrupa_Pagados_VF_Extracto(dbs, T_Final, T_Inicial)
dbs.Execute "SELECT MM_UDIS, TAXONOMIA, NR_R, BANCO, CSG, CSF, AGRUPAMIENTO, AGRUPAMIENTO_ID, " _
    & "CICLOS_RESCAT_VF AS CICLOS_RESCAT, TIPGAR_CLAVE, [ANIO(PAGO)], MPAGADO, " _
    & "ESQUEMA_VF AS ESQUEMA, [ANIO(PAGO)_New], Intermediario_Id, TPRO_CLAVE, ORDEN_PAGO " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "]; "
End Function

Function Une_Pagos_Recup(dbs, T_Final, T_Izquierda, T_Derecha)
dbs.Execute "SELECT A.*, B.Monto_Pagado_Mn/1000000 as [MPAGADO (MDP)] " _
          & "INTO [" & T_Final & "] " _
          & "FROM [" & T_Izquierda & "] as A LEFT JOIN [" & T_Derecha & "] as B " _
          & "ON A.CONCATENAR_SALDOS = B.CONCATENAR_SALDOS; "
End Function

Function Tabla_SevObs_Recup(dbs, T_Final, T_Inicial)
'Cambiar manual el año de recuperación el último año completo
'dbs.Close
'Set dbs = OpenDatabase("F:\INFO_NAFIN\GARANTIAS\Garantias\Curva_Severidad\2015\201503\Querie_CurvaRecuperada_201503.accdb")
dbs.Execute "SELECT A.MM_UDIS, A.TAXONOMIA, A.NR_R, A.BANCO, A.CSG, AGRUPAMIENTO, AGRUPAMIENTO_ID, A.TIPGAR_CLAVE, " _
    & "A.ESQUEMA, A.TPRO_CLAVE, A.[ANIO(PAGO)_New], 'Recup' as CONCEPTO, " _
    & "IIF(A.Anio_REG_RECUP<=2018,'Inicio','Acum') AS PERIODO, " _
    & "A.MRECUP_TOT AS MONTO " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "] as A " _
    & "WHERE A.CICLOS_RESCAT=1 AND A.[ANIO(PAGO)]<=2018 AND A.IND_ENTRA=1; "
End Function

Function Tabla_SevObs_Pagos(dbs, T_Final, T_Inicial)
'Cambiar manual al último año completo
'dbs.Close
'Set dbs = OpenDatabase("F:\INFO_NAFIN\GARANTIAS\Garantias\Curva_Severidad\2015\201502\Querie_CurvaRecuperada_201502.accdb")
dbs.Execute "SELECT B.MM_UDIS, B.TAXONOMIA, B.NR_R, B.BANCO, B.CSG, B.AGRUPAMIENTO, B.AGRUPAMIENTO_ID, B.TIPGAR_CLAVE, " _
    & "B.ESQUEMA, B.TPRO_CLAVE, B.[ANIO(PAGO)_New], 'Pago' as CONCEPTO, " _
    & "'Inicio' AS PERIODO, " _
    & "B.MPAGADO AS MONTO " _
    & "INTO [" & T_Final & "] " _
    & "FROM [" & T_Inicial & "] as B " _
    & "WHERE B.CICLOS_RESCAT=1 AND B.[ANIO(PAGO)]<=2018; "
End Function

Function Inserta_Filas_IN(dbs, BasePagos, TablaInicial, TablaFinal)
    dbs.Execute "insert into [" & TablaFinal & "] " _
        & "select * from [" & TablaInicial & "]" _
        & BasePagos & ";"
End Function

Function BorrarTabla(dbs, Tabla)
    dbs.Execute "DROP TABLE " & Tabla & ";"
End Function