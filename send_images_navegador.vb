Sub Generar_PDF()

For i = 1 To Cells(1, 22)

Cells(10, 5) = Cells(i + 1, 17)

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        Cells(2, 22) & "\" & Cells(i + 1, 19) & "\" & Format(Now, "yyyy_mm") & "_" & Cells(10, 5) & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
        
Next i

MsgBox "DETALLES REALIZADOS CON ÉXITO", vbInformation

End Sub
Sub auto_open()
UserForm1.Show
End Sub
Sub gallardo()
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Dim FileQUITA As String
Dim napoli As CDO.Message
Windows("Facturas_Agentes.xlsm").Activate
Sheets("Comprobante").Select
tesl = Range("aa1").Value
spax = Range("aa2").Value
Range("q1").Select
river = Selection.End(xlDown).Row - 1
For i = 1 To river
Set napoli = New CDO.Message
With napoli.Configuration.Fields
.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = tesl
.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = spax
.Update
End With
FileQUITA = Cells(2, 22) & "\" & Cells(i + 1, 19) & "\" & Format(Now, "yyyy_mm") & "_" & Cells(10, 5) & ".pdf"
para = Range("r" & i + 1)
asunto = Range("p" & i + 1)
With napoli
.Subject = asunto
.From = tesl
.To = para
.AddRelatedBodyPart "D:\pruebas\img\img.jpg", "img.jpg", 0
.HTMLBody = .HTMLBody & "<br><B>Señor agente tenga en cuenta lo siguiente:</B><br><p></p><p></p>" _
& "<center><img src='cid:img.jpg'></center>" _
& "<br>Cordialmente, <br>Gerencia de Comisiones</font></span>"
.AddAttachment FileQUITA
End With
napoli.Send
Set napoli = Nothing
Next i
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "DETALLES ENVIADOS CON ÉXITO", vbInformation
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
End Sub