Function Prepara_base()
' Eliminar todas las tablas y compactar la base

D = DateSerial(Year(Now()), Month(Now()) - 1, 1)
a = Format(D, "yyyy")
m = Format(D, "mm")
cierre = Format(D, "yyyymm")

SiExisteTabla_Borra LLAVE
SiExisteTabla_Borra PROGRAMA

'Vincula nuevos catálogos
    wd = "D:\DAR\proyecto_mejora_fotos\2. Nuevas fotos\"
    wd_external = wd & "data\external\"
    wd_processed = wd & "data\processed\"
    wd_processed_fotos = wd_processed & "Fotos\"
    wd_processed_fotos_cierre = wd_processed_fotos & cierre "\"

    db_catalogos = wd_external & "Catálogos_" & cierre & ".accdb"

    DoCmd.TransferDatabase acLink, "Microsoft Access", db_catalogos, acTable, "LLAVE", "LLAVE"
    DoCmd.TransferDatabase acLink, "Microsoft Access", db_catalogos, acTable, "PROGRAMA", "PROGRAMA"
    
'Importa nuevas fotos
    Ruta = "E:\Users\jhernandezr\DAR\garantias\fotos\202410\Fotos\202410\"
    
    F_R = wd_processed_fotos_cierre & "FotoRevolventesCohortes_" & cierre & "_VF.accdb"
    F_NR = wd_processed_fotos_cierre & "FotoSimplesCohortes_" & cierre & "_VF.accdb"
    DoCmd.TransferDatabase acImport, "Microsoft Access", F_R, acTable, "VF_Foto_R_" & cierre & "_VCohortes", "VF_Foto_R_" & cierre
    DoCmd.TransferDatabase acImport, "Microsoft Access", F_NR, acTable, "VF_Foto_NR_" & cierre, "VF_Foto_NR_" & cierre


'Elimina las columnas originales de la foto de revolventes
Dim sql1 As String
    sql1 = "UPDATE VF_Foto_R_" & cierre & " " & _
          "SET Estrato_Id_Original = Estrato_Id , TIPO_PERSONA_Original = TIPO_PERSONA "
DoCmd.RunSQL sql1

Dim sql2 As String
        sql2 = "alter table VF_Foto_R_" & cierre & " " & _
           "drop constraint Estrato_Id "
DoCmd.RunSQL sql2

Dim sql3 As String
        sql3 = "alter table VF_Foto_R_" & cierre & " " & _
           "drop column TIPO_PERSONA,Estrato_Id "
DoCmd.RunSQL sql3

End Function

