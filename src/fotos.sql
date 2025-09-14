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



--- Fotos Revolventes
--- Agrupa_x_Taxo
SELECT A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.NOMBRE, A.TAXONOMIA, A.CLAVE_TAXO, A.NR_R, A.INTER_CLAVE, A.CSG,
MIN(A.FECHA_VALOR1) as FECHA_VALOR1, MIN(A.FECHA_REGISTRO_GARANTIA) as FECHA_REGISTRO_GARANTIA,
SUM(A.[MGI (MDP)]) as [MGI (MDP)], AVG(A.PLAZO) as PLAZO, AVG(A.PLAZO_DIAS) as PLAZO_DIAS, MAX(A.FVTO) as FVTO,
Min(A.FECHA_PAGO) as FECHA_PAGO, SUM(A.INCUMPLIDO) as PAGADAS, SUM(A.[MPAGADO (MDP)]) as [MPAGADO (MDP)],
SUM(A.[MONTO CREDITO (MDP)]) as [MONTO CREDITO (MDP)], MIN(A.FECHA_VALOR) as FECHA_VALOR,
SUM(A.[SALDO (MDP)]) as [SALDO (MDP)], MIN(A.FECHA_REGISTRO1) as FECHA_REGISTRO1,
IIF(Min(A.FECHA_PRIMER_INCUM) IS NULL, cdate(Format('30/12/1899','dd/mm/yyyy')), Min(A.FECHA_PRIMER_INCUM)) as FECHA_PRIMER_INCUM,
MAX(A.MM_UDIS) as MM_UDIS, COUNT(A.NUM_GAR) as NUM_GAR, MAX(A.INCUMPLIDO) as INCUMPLIDO, First(A.ESQUEMA) as ESQUEMA,
SUM(A.[MONTOTOTAL (MDP)]) as [MONTOTOTAL (MDP)], SUM(A.[RECUPERADOS (MDP)]) as [TOT RECUP (MDP)], SUM(A.[RESCATADOS (MDP)]) as [TOT RESCAT (MDP)]
INTO [Microcredito]
FROM tbl_vf_foto_r as A
where (A.[TAXONOMIA]='GARANTIA MICROCREDITO')
GROUP BY A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.NOMBRE, A.TAXONOMIA, A.CLAVE_TAXO, A.NR_R, A.INTER_CLAVE, A.CSG 
ORDER BY A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.NOMBRE, A.TAXONOMIA, A.CLAVE_TAXO, A.NR_R, A.INTER_CLAVE, A.CSG;

-- Genera_x_Taxo
SELECT A.BANCO, A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, A.NOMBRE, A.TAXONOMIA, A.CLAVE_TAXO, A.NR_R, A.INTER_CLAVE, A.CSG, 
A.FECHA_VALOR1, A.FECHA_REGISTRO_GARANTIA, 
A.[MGI (MDP)], A.PLAZO, A.PLAZO_DIAS, A.FVTO, 
A.FECHA_PAGO, A.INCUMPLIDO as PAGADAS, A.[MPAGADO (MDP)], 
A.[MONTO CREDITO (MDP)],A.FECHA_VALOR, 
A.[SALDO (MDP)], A.FECHA_REGISTRO1, 
IIF(A.FECHA_PRIMER_INCUM IS NULL, cdate(Format('30/12/1899','dd/mm/yyyy')), A.FECHA_PRIMER_INCUM) as FECHA_PRIMER_INCUM, 
A.MM_UDIS, A.NUM_GAR, A.INCUMPLIDO, 
A.[MONTOTOTAL (MDP)], A.[RECUPERADOS (MDP)] as [TOT RECUP (MDP)], A.[RESCATADOS (MDP)] as [TOT RESCAT (MDP)] 
INTO [Empresarial_x_Disposicion] 
FROM [tbl_vf_foto_r] as A 
where (A.[TAXONOMIA]='GARANTIA EMPRESARIAL');

--- Agrega_Campos_x_Taxo
alter table Empresarial_x_Disposicion
add column [No_Acreditados_Saldo>0] Double;
update TaxoBase
set [No_Acreditados_Saldo>0] = IIF([SALDO (MDP)]>0,1,0);
set [Saldo^2] = [SALDO (MDP)]*[SALDO (MDP)];
set [Count] = IIF([SALDO (MDP)]>0,1,0);
set [ANTIG_CLIENTE_MESES] = IIF([SALDO (MDP)]>0,(YEAR(fecha_base)-YEAR(FECHA_VALOR1))*12+(MONTH(fecha_base)-MONTH(FECHA_VALOR1)),0);
set [ANTIG_CLIENTE_AÑOS] = IIF([SALDO (MDP)]>0,IIF(ANTIG_CLIENTE_MESES<=12,1,IIF(ANTIG_CLIENTE_MESES<=24,2,IIF(ANTIG_CLIENTE_MESES<=36,3,4))),0);
set [RESTANTE_MESES] = IIF((IIF([SALDO (MDP)]>0,(YEAR(FVTO)-YEAR(fecha_base))*12+(MONTH(FVTO)-MONTH(fecha_base)),0))<0,0,(IIF([SALDO (MDP)]>0,(YEAR(FVTO)-YEAR(fecha_base))*12+(MONTH(FVTO)-MONTH(fecha_base)),0)));
set [RESTANTE_POND] = RESTANTE_MESES*[SALDO (MDP)];
set [VIGENTES] = IIF(FVTO+180> fecha_base,1,0);
set [REMANENTE_MESES] = IIF([SALDO (MDP)]>0,IIF(FVTO> fecha_base,(YEAR(FVTO)-YEAR(fecha_base))*12+(MONTH(FVTO)-MONTH(fecha_base)),0),0);
set [REMANENTE_AÑOS] = IIF(REMANENTE_MESES<=12,1,IIF(REMANENTE_MESES<=24,2,IIF(REMANENTE_MESES<=36,3,IIF(REMANENTE_MESES<=48,4,5))));
set [REMANENTE_MESES+180] = IIF([SALDO (MDP)]>0,IIF(FVTO+180>fecha_base,(YEAR(FVTO)-YEAR(fecha_base))*12+(MONTH(FVTO)-MONTH(fecha_base))+180/30,0),0);
set [REMANENTE_AÑOS+180] = IIF(REMANENTE_MESES+180<=12,1,IIF(REMANENTE_MESES+180<=24,2,IIF(REMANENTE_MESES+180<=36,3,IIF(REMANENTE_MESES+180<=48,4,5))));
set [Antig_Cliente_Meses_Pond] = ANTIG_CLIENTE_MESES * [SALDO (MDP)];
set [RESTANTE_DIAS] = IIF([SALDO (MDP)]>0,(FVTO-fecha_base),0);
set [RESTANTE_DIAS_POND] = RESTANTE_DIAS*[SALDO (MDP)];




--- Fotos Simples
--- Base_Simple_x_Taxo
SELECT
A.CLAVE_CREDITO, A.FECHA_VALOR1, A.TIPO_PERSONA, A.NOMBRE, A.RFC, A.FECHA_REGISTRO_GARANTIA, A.[MGI (MDP)], 
A.PLAZO, A.PLAZO_DIAS, A.FVTO, A.BANCO, A.FECHA_PAGO, A.INCUMPLIDO as PAGADAS, A.[MPAGADO (MDP)], A.FECHA_REGISTRO1,  
A.[MONTO CREDITO (MDP)], A.FECHA_VALOR, A.INTER_CLAVE, A.TPRO_CLAVE, A.NR_R, A.CSG, A.[SALDO (MDP)], 
IIF(A.FECHA_PRIMER_INCUM IS NULL, cdate(Format('30/12/1899','dd/mm/yyyy')),A.FECHA_PRIMER_INCUM) as FECHA_PRIMER_INCUM, 
A.CLAVE_TAXO, A.TAXONOMIA, A.MM_UDIS, IIF(A.CLAVE_CREDITO IS NULL,0,1) as NUM_GAR, A.INCUMPLIDO, A.ESQUEMA, 
A.[MONTOTOTAL (MDP)], A.[RECUPERADOS (MDP)], A.[RESCATADOS (MDP)], A.AGRUPAMIENTO, A.AGRUPAMIENTO_ID, 
A.PORCENTAJE_GARANTIZADO, A.PLAZO_BUCKET, A.Programa_Original, A.Programa_Id 
INTO [Microcredito] 
FROM [tbl_vf_foto_nr] as A 
where (A.[TAXONOMIA]='GARANTIA MICROCREDITO');


-- Campos Extras
set MGI_VIVOS = (IIF(FVTO+180>fecha_base,1,0)*[MGI (MDP)]);
set MGI_MALOS_VIVOS = (IIF(FVTO+180>fecha_base,1,0)*[MGI (MDP)]*INCUMPLIDO);
set MPAGADO_VIVOS = (IIF(FVTO+180>fecha_base,1,0)*[MPAGADO (MDP)]);
set MRECUP_VIVOS = (IIF(FVTO+180>fecha_base,1,0)*[RECUPERADOS (MDP)]);
set MGI_CAD = (IIF(FVTO+180<=fecha_base,1,0)*[MGI (MDP)]);
set MGI_MALOS_CAD = (IIF(FVTO+180<=fecha_base,1,0)*[MGI (MDP)]*INCUMPLIDO);
set MPAGADO_CAD = (IIF(FVTO+180<=fecha_base,1,0)*[MPAGADO (MDP)]);
set MRECUP_CAD = (IIF(FVTO+180<=fecha_base,1,0)*[RECUPERADOS (MDP)]);
set SALDO_VIVOS = (IIF(FVTO+180>fecha_base,1,0)*[SALDO (MDP)]);
set SALDO_CADUCOS = (IIF(FVTO+180<=fecha_base,1,0)*[SALDO (MDP)]);
set MGI_VIVOS^2 = [MGI_VIVOS]*[MGI_VIVOS];
set #VIVAS = IIF(FVTO+180>fecha_base,1,0);
set MGI_INCMPL = [MGI (MDP)]*INCUMPLIDO;
set Count = 1;
set FECHA_PAGO1 = IIF(FECHA_PAGO=0,NULL, cdate(Format(dateserial(Year(FECHA_PAGO),Month(FECHA_PAGO),'01'),'dd/mm/yyyy')));
set AñoOtor = [AñoOtor] = YEAR(FECHA_VALOR1);
set PTRANSCURRIDO = (IIF(1>((fecha_base - FECHA_VALOR)/(FVTO-FECHA_VALOR+180)),((fecha_base - FECHA_VALOR)/(FVTO-FECHA_VALOR+180)),1));
set PTRANS_PON = [PTRANSCURRIDO]*[MGI (MDP)];
set Semestre = IIF(MONTH(FECHA_VALOR)<7,cdate(Format(dateserial(YEAR(FECHA_VALOR),'01','01'),'dd/mm/yyyy')),cdate(Format(dateserial(YEAR(FECHA_VALOR),'02','01'),'dd/mm/yyyy')));
set MESES_REM_POND = (1-PTRANSCURRIDO)*((YEAR(FVTO)-YEAR(FECHA_VALOR1))*12+(MONTH(FVTO)-MONTH(FECHA_VALOR1))+6)*[SALDO (MDP)];
set Con_Saldo = IIF([SALDO (MDP)]>0,1,0);
set RESTANTE_MESES = IIF(0<(IIF([Con_Saldo]=1,(YEAR(FVTO)-YEAR(fecha_base))*12+(MONTH(FVTO)-MONTH(fecha_base)),0)),(IIF([Con_Saldo]=1,(YEAR(FVTO)-YEAR(fecha_base))*12+(MONTH(FVTO)-MONTH(fecha_base)),0)),0);
set RESTANTE_POND = RESTANTE_MESES*[SALDO (MDP)];
set SALDO^2 = [SALDO (MDP)] * [SALDO (MDP)];
set RESTANTE_DIAS = IIF([SALDO (MDP)]>0,(FVTO-fecha_base),0);
set RESTANTE_DIAS_POND = RESTANTE_DIAS*[SALDO (MDP)];




--- Fotos_VF Ajuste fotos 2
SELECT 
    [BUCKET], [CAMBIO], [MCrédito_MM_UDIS], [MM_UDIS], 
    fotos_recup.[INTER_CLAVE], [NOMBRE], [RFC], [TIPO_PERSONA], 
    fotos_recup.[CLAVE_CREDITO], [FECHA_VALOR], [PLAZO_DIAS], 
    [PLAZO], [FVTO], [FECHA_REGISTRO_GARANTIA], [MGI (MDP)], 
    [PORCENTAJE_GARANTIZADO], [BANCO], [FECHA_PRIMER_INCUM], 
    [MONTO CREDITO (MDP)], [SALDO (MDP)], [TPRO_CLAVE], [CLAVE_TAXO], 
    [TAXONOMIA], [NR_R], [FECHA_VALOR1], [FECHA_REGISTRO1], [NUM_GAR], 
    [CSG], [PLAZO_BUCKET], pagos_agrup.pagos_aux AS [MPAGADO (MDP)], 
    [PAGADAS], [INCUMPLIDO], [FECHA_PAGO], [Programa_Original], [Programa_Id], 
    [Estrato_Id], [Sector_Id], [Estado_Id], [Tipo_Credito_Id], 
    [Porcentaje_Comision_Garantia], [Tasa_Id], [Tasa_Interes], 
    [MGI (MDP) Original], [AGRUPAMIENTO_ID], [ESQUEMA], [SUBESQUEMA], 
    [AGRUPAMIENTO], [FONDOS_CONTRAGARANTIA], [CONREC_CLAVE], [Describe_Desrec], 
    [MONTOTOTAL (MDP)], fotos_recup.recup_aux AS [RECUPERADOS (MDP)], 
    [RESCATADOS (MDP)], fecha_recup, monedaux AS Moneda_Id, pago_or, recup_or 
INTO Foto_VF
FROM (SELECT * FROM (SELECT * FROM VF_Foto_NR UNION ALL SELECT * FROM VF_Foto_R)  AS fotos 
        LEFT JOIN (SELECT Intermediario_Id, Numero_Credito, sum(Monto_Total_Mn)/1000000 AS recup_aux, 
            max(Fecha) AS fecha_recup, sum(Monto_Total_Mn/tipo_cambio_cierre)/1000000 AS recup_or 
            FROM Recupera_con_Pagos_Flujos 
            WHERE Estatus='CR' or Estatus="D" or Estatus="E" or Estatus="RAC" or Estatus="RAR" or Estatus="RI" 
            GROUP BY Intermediario_Id, Numero_Credito)  AS recup_agrup 
            ON (fotos.INTER_CLAVE=recup_agrup.Intermediario_Id) AND (fotos.CLAVE_CREDITO=recup_agrup.Numero_Credito))
              AS fotos_recup 
              LEFT JOIN (SELECT Intermediario_Id, Numero_Credito, sum(Monto_Pagado_Mn)/1000000 AS pagos_aux, 
              max(moneda_id) AS monedaux, sum(Monto_Pagado_Mn/TC)/1000000 AS pago_or 
              FROM Pagadas_Global_VF GROUP BY Intermediario_Id, Numero_Credito)  AS pagos_agrup ON (fotos_recup.INTER_CLAVE=pagos_agrup.Intermediario_Id) AND (fotos_recup.CLAVE_CREDITO=pagos_agrup.Numero_Credito);
