Sub Principal()

MesNum = 202411
Anio = Left(MesNum, 4)
Mes1 = IIf(Mid(MesNum, 5, 1) = 0, Right(MesNum, 1), Right(MesNum, 2))
mes = Mes_Letra(Mes1) & Mid(MesNum, 3, 2)


wd = "E:\Users\jhernandezr\DAR\garantias\reporte\fotos\"
wd_external = wd & "data\external\"
wd_processed = wd & "data\processed\"
wd_processed_dwh = wd_processed & "DWH\"
wd_processed_dwh_bases_finales = wd_processed_dwh & "bases_finales\"
wd_processed_dwh_entregables = wd_processed_dwh & "Entregables\"
wd_processed_fotos = wd_processed & "Fotos\"
wd_processed_curvarecup = wd_processed & "CurvaRecup\"
wd_raw = wd & "data\raw\"
wd_staging = wd & "data\staging\"
wd_validations = wd & "data\validations\"


'Inputs
db_catalogos = wd_external & "Catálogos_" & mes & ".accdb"
db_base = wd_staging & "Base.accdb"
db_querie_pagadas = wd_processed_dwh & "Querie_Pagadas_" & MesNum & ".accdb"
db_querie_union_flujos = wd_processed_dwh & "Querie_UnionFlujos_" & MesNum & ".accdb"

'Outputs
db_querie_curva_recuperada = wd_processed_curvarecup & "Querie_CurvaRecuperada_" & MesNum & ".accdb"
db_querie_curva_recuperada_aux1 = wd_processed_curvarecup & "Querie_CurvaRecuperada_" & MesNum & "_AUX1.accdb"
db_querie_curva_recuperada_aux2 = wd_processed_curvarecup & "Querie_CurvaRecuperada_" & MesNum & "_AUX2.accdb"

'Inicializo todo
tbl_recupera_con_pagos_flujos_sbancomext = "Recupera_con_Pagos_Flujos_" & MesNum & "_Ord_sBancomext"
tbl_recuperadas_detalle = "Recuperadas_Detalle_" & MesNum
tbl_recuperadas_detalle_cllave = "Recuperadas_Detalle_" & MesNum & "_cLlave"
tbl_pagadas_global_vf_sbancomext = "Pagadas_Global_VF_" & MesNum & "_sBancomext"
tbl_pagadas_detalle = "Pagadas_Detalle_" & MesNum
tbl_pagadas_detalla_cllave = "Pagadas_Detalle_" & MesNum & "_cLlave"
tbl_llave_ = "Llave_" & MesNum
tbl_llave = "LLAVE"
tbl_recup_previo = "RECUP_PREVIO_" & MesNum
tbl_ciclos_rescate = "CICLOS_RESCATE"
tbl_curva_recup = "CURVA_RECUP_" & MesNum
tbl_recup_agrup = "RECUP_AGRUP_" & MesNum
tbl_recup_agrup_inter = tbl_recup_agrup & "_inter"
tbl_recup_agrup_completa = tbl_recup_agrup & "_Completa"
tbl_utl_pago = "TABLA_ULT_PAGO"
tbl_recup_previo_cup = "RECUP_PREVIO_" & MesNum & "_cUP"
tbl_pagadas_previo = "PAGADAS_PREVIO_" & MesNum
tbl_pagadas_detalle_vf = "PAGADAS_DETALLE_VF_" & MesNum
tbl_pagadas_agrup = "PAGADAS_AGRUP_" & MesNum
tbl_pagadas_agrup_inter = tbl_pagadas_agrup & "_inter"
tbl_pagadas_agrup_completa = tbl_pagadas_agrup & "_Completa"
tbl_curva_recup_vf = "CURVA_RECUP_" & MesNum & "_VF"
tbl_sev_obs = "TABLA_SEV_OBS_" & MesNum
tbl_sev_obs_temp = "TABLA_SEV_OBS_" & MesNum & "_Temp"


If ExisteRuta(wd_processed_curvarecup) = False Then
    MkDir (wd_processed_curvarecup)
End If
'Crea Base Curva Recuperadas
If ExisteRuta(db_querie_curva_recuperada) = False Then
    Copia = Application.CompactRepair(db_base, wd_processed_curvarecup & "Copia de Seguridad.accdb")
    Name wd_processed_curvarecup & "Copia de Seguridad.accdb" As db_querie_curva_recuperada
End If
'Importar Recuperadas (El detalle de las recuperaciones)

Call Vincula_Tabla(db_querie_union_flujos, db_querie_curva_recuperada, tbl_recupera_con_pagos_flujos_sbancomext, tbl_recuperadas_detalle)
Call Vincula_Tabla(db_catalogos, db_querie_curva_recuperada, "Estatus", "Estatus")
'Importar Pagadas (El detalle de los pagos)

Call Vincula_Tabla(db_querie_pagadas, db_querie_curva_recuperada, tbl_pagadas_global_vf_sbancomext, tbl_pagadas_detalle)
'Importar Llave de Javier
'R_Llave = Unidad_Gris & ":\INFO_NAFIN\GARANTIAS\Garantias\CURVA_SEVERIDAD\" & Anio & "\"
'BDE_Llave = "Llave_201312.xlsx"
Call Vincula_Tabla(db_catalogos, db_querie_curva_recuperada, tbl_llave, tbl_llave_)
'Call Importar_Hoja_Excel(R_Llave & BDE_Llave, db_querie_curva_recuperada, tbl_llave_)
'---------------------------------------------------------------------------------
Set dbs = OpenDatabase(db_querie_curva_recuperada)
    'Cruza con Llave y AgrupamientosLlave
    Call Cruza_Llave(dbs, tbl_recuperadas_detalle_cllave, tbl_recuperadas_detalle, tbl_llave_)
    Call Cruza_Llave(dbs, tbl_pagadas_detalla_cllave, tbl_pagadas_detalle, tbl_llave_)
    'Recuperaciones
    Call Base_Recup_Detalle(dbs, tbl_recup_previo, tbl_recuperadas_detalle_cllave, "Estatus") 'todavía trae el CSF
    
    Call CiclosRescate(dbs, tbl_ciclos_rescate, tbl_recup_previo)
    
    'ACOV 201805 proceso para automatizar que la base ya está muy pesada
    dbs.Close
    'Compacto la base en otra base llamada Aux1
    DAO.DBEngine.CompactDatabase db_querie_curva_recuperada, db_querie_curva_recuperada_aux1
    'Elimino la base pesada
    Kill db_querie_curva_recuperada
    'Vuelvo a crear la base
    Copia = Application.CompactRepair(db_base, wd_processed_curvarecup & "Copia de Seguridad.accdb")
    Name wd_processed_curvarecup & "Copia de Seguridad.accdb" As db_querie_curva_recuperada
    'Vinculo las tablas de la otra base
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_ciclos_rescate, tbl_ciclos_rescate)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_recup_previo, tbl_recup_previo)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_recuperadas_detalle_cllave, tbl_recuperadas_detalle_cllave)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_pagadas_detalla_cllave, tbl_pagadas_detalla_cllave)
    'Vinculo las del catálogo
    Call Vincula_Tabla(db_catalogos, db_querie_curva_recuperada, tbl_llave, tbl_llave_)
    Call Vincula_Tabla(db_catalogos, db_querie_curva_recuperada, "Estatus", "Estatus")
    Set dbs = OpenDatabase(db_querie_curva_recuperada)
    
    'Continúo con el original
    
    Call Cruza_CiclosRescate(dbs, tbl_curva_recup, tbl_recup_previo, tbl_ciclos_rescate) 'todavía trae el CSF
    Call AgrupaRecup(dbs, tbl_recup_agrup_inter, tbl_curva_recup, "", "") 'OJO
    Call AgrupaRecup(dbs, tbl_recup_agrup_inter & "_Mes", tbl_curva_recup, "Month(Fecha_Registro) as Mes_REG_RECUP, ", "Month(Fecha_Registro), ")
    Call Agrupa_Recup_VF_Completo(dbs, tbl_recup_agrup_completa & "_Mes", tbl_recup_agrup_inter & "_Mes") 'OJO
    Call Agrupa_Recup_VF_Completo(dbs, tbl_recup_agrup_completa, tbl_recup_agrup_inter) 'OJO
    Call Agrupa_Recup_VF_Extracto(dbs, tbl_recup_agrup, tbl_recup_agrup_completa) 'OJO
    'Pagos
    Call Agrupa_UltimoPago(dbs, tbl_utl_pago, tbl_pagadas_detalla_cllave)
    Call Toma_UltimoPago(dbs, tbl_recup_previo_cup, tbl_pagadas_detalla_cllave, tbl_utl_pago)
    Call Base_Pagadas_Detalle(dbs, tbl_pagadas_previo, tbl_recup_previo_cup)
   
    dbs.Close
    Set dbs = OpenDatabase(db_querie_curva_recuperada)
    
    Call Cruza_CiclosRescate(dbs, tbl_pagadas_detalle_vf, tbl_pagadas_previo, tbl_ciclos_rescate)
    Call Agrupa_Pagos(dbs, tbl_pagadas_agrup_inter, tbl_pagadas_detalle_vf) ' OJO
    Call Agrupa_Pagados_VF_Completo(dbs, tbl_pagadas_agrup_completa, tbl_pagadas_agrup_inter) 'OJO
    Call Agrupa_Pagados_VF_Extracto(dbs, tbl_pagadas_agrup, tbl_pagadas_agrup_completa) 'OJO
    'Une Recuperaciones y Pagos
    
    'ACOV 201805 proceso para automatizar que la base ya está muy pesada
    dbs.Close
    'Compacto la base en otra base llamada Aux1
    DAO.DBEngine.CompactDatabase db_querie_curva_recuperada, db_querie_curva_recuperada_aux2
    'Elimino la base pesada
    Kill db_querie_curva_recuperada
    'Vuelvo a crear la base
    Copia = Application.CompactRepair(db_base, wd_processed_curvarecup & "Copia de Seguridad.accdb")
    Name wd_processed_curvarecup & "Copia de Seguridad.accdb" As db_querie_curva_recuperada
    'Vinculo las tablas de la base AUX1
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_ciclos_rescate, tbl_ciclos_rescate)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_recup_previo, tbl_recup_previo)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_recuperadas_detalle_cllave, tbl_recuperadas_detalle_cllave)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux1, db_querie_curva_recuperada, tbl_pagadas_detalla_cllave, tbl_pagadas_detalla_cllave)
    'Vinculo las del catálogo
    Call Vincula_Tabla(db_catalogos, db_querie_curva_recuperada, tbl_llave, tbl_llave_)
    Call Vincula_Tabla(db_catalogos, db_querie_curva_recuperada, "Estatus", "Estatus")
    'Vinculo las de AUX1
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_curva_recup, tbl_curva_recup)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_pagadas_agrup, tbl_pagadas_agrup)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_pagadas_agrup_completa, tbl_pagadas_agrup_completa)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_pagadas_agrup_inter, tbl_pagadas_agrup_inter)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_pagadas_detalle_vf, tbl_pagadas_detalle_vf)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_pagadas_previo, tbl_pagadas_previo)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_recup_agrup, tbl_recup_agrup)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_recup_agrup_completa, tbl_recup_agrup_completa)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_recup_agrup_inter, tbl_recup_agrup_inter)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_utl_pago, tbl_utl_pago)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_recup_previo_cup, tbl_recup_previo_cup)
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_recup_agrup_completa & "_Mes", tbl_recup_agrup_completa & "_Mes")
    Call Vincula_Tabla(db_querie_curva_recuperada_aux2, db_querie_curva_recuperada, tbl_recup_agrup_inter & "_Mes", tbl_recup_agrup_inter & "_Mes")
    Set dbs = OpenDatabase(db_querie_curva_recuperada)
   
   'Proceso original
    Call Une_Pagos_Recup(dbs, tbl_curva_recup_vf, tbl_curva_recup, tbl_pagadas_detalle_vf)
    
    'Severidad Observada
    Call Tabla_SevObs_Recup(dbs, tbl_sev_obs, tbl_recup_agrup)
    Call Tabla_SevObs_Pagos(dbs, tbl_sev_obs_temp, tbl_pagadas_agrup)
    Call Inserta_Filas_IN(dbs, "", tbl_sev_obs_temp, tbl_sev_obs)
    Call BorrarTabla(dbs, tbl_sev_obs_temp)
    
    'ACOV 201902 Tablas de entrega para Daf/Jera
    dbs.Execute "SELECT * INTO " & tbl_pagadas_agrup & "_VF FROM " & tbl_pagadas_agrup & " WHERE TAXONOMIA NOT IN ('GARANTIAS BANCOMEXT','GARANTIAS SHF/LI FINANCIERO','GARANTIAS BANSEFI'); "
    dbs.Execute "SELECT * INTO " & tbl_recup_agrup & "_VF FROM " & tbl_recup_agrup & " WHERE TAXONOMIA NOT IN ('GARANTIAS BANCOMEXT','GARANTIAS SHF/LI FINANCIERO','GARANTIAS BANSEFI'); "
    
    'ACOV 201902 crea consultas de validación de saldos
    dbs.Execute "SELECT TAXONOMIA, SUM(MPAGADO) AS S_MPagado INTO Valida_Pagos FROM " & tbl_pagadas_agrup & "_VF GROUP BY TAXONOMIA; "
    dbs.Execute "SELECT TAXONOMIA, SUM(MRECUP_TOT) AS S_MRcup, SUM(MRESCAT_TOT) AS S_MRescat INTO Valida_Recup_Rescat FROM " & tbl_recup_agrup & "_VF WHERE IND_ENTRA = 1 GROUP BY TAXONOMIA; "
    'FIDE
    dbs.Execute "SELECT TAXONOMIA, SUM(MPAGADO) AS S_MPagado INTO Valida_Pagos_FIDE FROM " & tbl_pagadas_agrup & "_VF WHERE AGRUPAMIENTO = 'FIDE' GROUP BY TAXONOMIA; "
    dbs.Execute "SELECT TAXONOMIA, SUM(MRECUP_TOT) AS S_MRcup, SUM(MRESCAT_TOT) AS S_MRescat INTO Valida_Recup_Rescat_FIDE FROM " & tbl_recup_agrup & "_VF WHERE AGRUPAMIENTO = 'FIDE' AND IND_ENTRA=1  GROUP BY TAXONOMIA; "
    
dbs.Close

If ExisteRuta(wd_processed_curvarecup & "SevObs_Meta_" & MesNum) = False Then
    MkDir (wd_processed_curvarecup & "SevObs_Meta_" & MesNum)
End If

Call Exporta_Access_a_Excel(wd_processed_curvarecup & tbl_pagadas_agrup & ".xlsx", db_querie_curva_recuperada, tbl_pagadas_agrup & "_VF")
Call Exporta_Access_a_Excel(wd_processed_curvarecup & tbl_recup_agrup & ".xlsx", db_querie_curva_recuperada, tbl_recup_agrup & "_VF")
Call Exporta_Access_a_Excel(wd_processed_curvarecup & "SevObs_Meta_" & MesNum & "\" & tbl_sev_obs & ".xlsx", db_querie_curva_recuperada, tbl_sev_obs)

End Sub