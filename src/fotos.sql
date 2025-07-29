SELECT A.*, (IIF(B.[Monto_Desembolso_Mn] is NULL, 0,B.[Monto_Desembolso_Mn])+IIF(B.[Interes_Desembolso_Mn] is NULL, 0,B.[Interes_Desembolso_Mn])+IIF(B.[Interes_Moratorios_Mn] is NULL, 0,B.[Interes_Moratorios_Mn]))/1000000 as [MPAGADO (MDP)],
IIF((IIF(B.[Monto_Desembolso_Mn] is NULL, 0,B.[Monto_Desembolso_Mn])+IIF(B.[Interes_Desembolso_Mn] is NULL, 0,B.[Interes_Desembolso_Mn])+IIF(B.[Interes_Moratorios_Mn] is NULL, 0,B.[Interes_Moratorios_Mn]))>0,1,0) as PAGADAS,
IIF((IIF(B.[Monto_Desembolso_Mn] is NULL, 0,B.[Monto_Desembolso_Mn])+IIF(B.[Interes_Desembolso_Mn] is NULL, 0,B.[Interes_Desembolso_Mn])+IIF(B.[Interes_Moratorios_Mn] is NULL, 0,B.[Interes_Moratorios_Mn]))>0,1,0) as INCUMPLIDO,
IIF(B.[Fecha_Garantia_Honrada] is NULL,cdate(Format('30/12/1899','dd/mm/yyyy')), B.[Fecha_Garantia_Honrada]) as FECHA_PAGO
INTO [Temp_BD_DWH_R_202506] 
FROM [BD_DWH_R_202506] A LEFT JOIN [PAGADAS_DETALLE_VF_202506] B 
on (cstr(A.Intermediario_Id)=cstr(B.Intermediario_Id) AND A.Numero_Credito=B.Numero_Credito)
IN BaseOrigen;


-- Z3
TRO_Temp = Recuperadas_Global_VF_202506_Origen_Temp
SELECT N.*, N.Numero_Credito & N.Intermediario_Id as Concatenado
into [" & Recuperadas_Global_VF_202506_Origen_Temp & "] 
FROM [" & Recuperadas_Global_VF_202506 & "] as N 

SELECT A.Numero_Credito, A.Intermediario_Id, A.NR_R, A.Producto,
IIF (A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]), 1, 0) AS ENTRA_RECUP,

IIF ((A.Fecha > 
    IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]) 
            AND (A.Estatus='D' or A.Estatus='E' or A.Estatus='RI' or A.Estatus='CR' or A.Estatus='RAR' or A.Estatus='RAC' or A.Estatus='CJ' or A.Estatus='CS' or A.Estatus='R' or A.Estatus='RJ' or A.Estatus='RS')), (nz(A.Monto_Mn ,0)+nz(A.Interes_Mn,0)+nz(A.Moratorios_Mn,0)+nz(A.Excedente_Mn,0)-nz(A.[Gastos_Juicio_Mn],0))/1000000, 0) AS [MONTOTOTAL (MDP)],

IIF ((A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]) AND (A.Estatus='D' or A.Estatus='E' or A.Estatus='RI' or A.Estatus='CR' or A.Estatus='RAR' or A.Estatus='RAC')), (nz(A.Monto_Mn,0)+nz(A.Interes_Mn,0)+nz(A.Moratorios_Mn,0)+nz(A.Excedente_Mn,0)-nz(A.[Gastos_Juicio_Mn],0))/1000000,0) AS [RECUPERADOS (MDP)],
IIF ((A.Fecha > IIF(B.[FECHA_PAGO] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),B.[FECHA_PAGO]) AND (A.Estatus='CJ' or A.Estatus='CS' or A.Estatus='R' or A.Estatus='RJ' or A.Estatus='RS')), (nz(A.Monto_Mn,0)+nz(A.Interes_Mn,0)+nz(A.Moratorios_Mn,0)+nz(A.Excedente_Mn,0)-nz(A.[Gastos_Juicio_Mn],0))/1000000,0) AS [RESCATADOS (MDP)]
INTO [Z3_RECUPCOHOR]
FROM [Recuperadas_Global_VF_202506_Origen_Temp] as A LEFT JOIN (SELECT M.*,  M.[Numero_Credito] & M.Intermediario_Id as Concatenado2 FROM [BD_DWH_R_202506] M) as B
on (A.Concatenado=B.Concatenado2)


SELECT  A.BUCKET, A.CAMBIO, A.[Monto _Credito_Mn]*A.CAMBIO AS MCrédito_MM_UDIS, A.[MM_UDIS], 
A.[Intermediario_Id] as INTER_CLAVE, A.Nombre_v1 as NOMBRE, A.[RFC Empresa / Acreditado] as RFC, A.[TIPO_PERSONA] as TIPO_PERSONA, A.[Numero_Credito] as CLAVE_CREDITO, 
A.[Fecha de Apertura] as FECHA_VALOR, IIF(A.[Plazo Días] IS NULL,0,A.[Plazo Días]) as PLAZO_DIAS, A.[Plazo] as PLAZO, A.[FVTO_Riesgosd] as FVTO, A.[Fecha Registro Alta] as FECHA_REGISTRO_GARANTIA, 
A.[Monto_Garantizado_Mn]/1000000 as [MGI (MDP)], A.[Porcentaje Garantizado] as PORCENTAJE_GARANTIZADO, A.[Razón Social (Intermediario)] as BANCO, IIF(A.[Fecha_Primer_Incumplimiento] IS NULL,cdate(Format('30/12/1899','dd/mm/yyyy')),A.[Fecha_Primer_Incumplimiento]) as FECHA_PRIMER_INCUM, 
A.[Monto _Credito_Mn]/1000000 as [MONTO CREDITO (MDP)], A.[Saldo_Contingente_Mn]/1000000 as [SALDO (MDP)], A.[TPRO_CLAVE] as TPRO_CLAVE, 
A.[Producto ID] as CLAVE_TAXO, A.[Producto] as TAXONOMIA, A.[NR_R], 
IIF(A.[Fecha de Apertura]=0,NULL, cdate(Format(dateserial(Year(A.[Fecha de Apertura]),Month(A.[Fecha de Apertura]),'01'),'dd/mm/yyyy'))) AS FECHA_VALOR1, 
IIF(A.[Fecha Registro Alta]=0,NULL, cdate(Format(dateserial(Year(A.[Fecha Registro Alta]),Month(A.[Fecha Registro Alta]),'01'),'dd/mm/yyyy'))) AS FECHA_REGISTRO1, 
IIF(A.[Numero_Credito] is NULL, 0,1) AS NUM_GAR, A.[CSG], 
IIF(Plazo<=12,1,IIF(Plazo<=24,2,IIF(Plazo<=36,3,4))) AS PLAZO_BUCKET, A.[MPAGADO (MDP)], A.PAGADAS, A.INCUMPLIDO, A.FECHA_PAGO, 
A.[Programa_Original] as Programa_Original, A.[Programa_Id] as Programa_Id, A.[Estrato_Id] as Estrato_Id, A.[Sector_Id] as Sector_Id, A.[Estado_Id] as Estado_Id, A.[Tipo_Credito_Id] as Tipo_Credito_Id, A.[Porcentaje de Comisión Garantia] as Porcentaje_Comision_Garantia, 
A.[Tasa_Id] as Tasa_Id, A.[Valor_Tasa_Interes] as [Tasa_Interes],  A.[Monto_Garantizado_Mn_Original]/1000000 as [MGI (MDP) Original], A.[AGRUPAMIENTO_ID], 
A.ESQUEMA, A.SUBESQUEMA, A.AGRUPAMIENTO, A.FONDOS_CONTRAGARANTIA, A.CONREC_CLAVE, A.Describe_Desrec 
INTO VF_Pagadas_R_202506 
FROM [BD_DWH_R_202506 A;