Sub marcolis()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Windows("PLANTILLA.xlsm").Activate
Sheets("RESIDUAL").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("maestra").Select
Range("A2").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("RESIDUAL").Select
Range("A2").Select
Sheets("CORRESPONSALIA").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("maestra").Select
Range("A2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("CORRESPONSALIA").Select
Range("A2").Select
Sheets("UP SELLING").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("maestra").Select
Range("A2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("UP SELLING").Select
Range("A2").Select
Sheets("RECONOCIMIENTO LOGISTICO").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("maestra").Select
Range("A2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("RECONOCIMIENTO LOGISTICO").Select
Range("A2").Select
Sheets("CPS").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("maestra").Select
Range("A2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("CPS").Select
Range("A2").Select
Sheets("CLARO UP").Select
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("maestra").Select
Range("A2").Select
Selection.End(xlDown).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Sheets("CLARO UP").Select
Range("A2").Select
Sheets("maestra").Select
Range("A1").Select
core = Selection.End(xlDown).Row
Range(Selection, Selection.End(xlDown)).Select
ActiveSheet.Range("$A$1:$A$" & core).RemoveDuplicates Columns:=1, Header:=xlYes
Range("A1").Select
core = Selection.End(xlDown).Row
Range("A2").Select
ActiveCell.Offset(0, 2).Select
ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-2],""_Detalle_liquidación_Claro Colombia_Enero P1 2022"")"
ActiveCell.Copy
Range("C2:C" & core).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select
MsgBox "LISTO PRIMERA FASE", vbInformation
Sheets("maestra").Select
Range("a1").Select
river = Selection.End(xlDown).Row - 1
Range("A2").Select
For i = 1 To river
valor = ActiveCell.Value
Selection.End(xlToRight).Select
nomlib = ActiveCell.Value
ss = nomlib & ".xlsx"
ActiveCell.Offset(1, 0).Select
Selection.End(xlToLeft).Select
Workbooks.Add.Worksheets.Add
Sheets("Hoja1").Select
Sheets("Hoja1").Name = "RESIDUAL_"
ChDir "D:\ARCHIVOS_AGENTES\ARCHIVOS"
ActiveWorkbook.SaveAs ("D:\ARCHIVOS_AGENTES\ARCHIVOS\" & nomlib & ".xlsx")
Windows("PLANTILLA.xlsm").Activate
Sheets("RESIDUAL").Select
Set RangoDatos = Sheets("RESIDUAL").UsedRange
RangoDatos.AutoFilter Field:=1, Criteria1:=valor
ultima = Sheets("RESIDUAL").Range("A" & Rows.Count).End(xlUp).Row
Sheets("RESIDUAL").Range("A1:N" & ultima).Copy
Windows(ss).Activate
Sheets("RESIDUAL_").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("G1").Select
Selection.Columns.ColumnWidth = 20
Range("M1").Select
Selection.Columns.ColumnWidth = 20
Sheets.Add After:=Sheets(Sheets.Count)
Sheets("Hoja2").Select
Sheets("Hoja2").Name = "CORRESPONSALIA_"
Windows("PLANTILLA.xlsm").Activate
Sheets("CORRESPONSALIA").Select
Set RangoDatos = Sheets("CORRESPONSALIA").UsedRange
RangoDatos.AutoFilter Field:=1, Criteria1:=valor
ultima = Sheets("CORRESPONSALIA").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CORRESPONSALIA").Range("A1:S" & ultima).Copy
Windows(ss).Activate
Sheets("CORRESPONSALIA_").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("E1").Select
Selection.Columns.ColumnWidth = 20
Range("L1").Select
Selection.Columns.ColumnWidth = 30
Range("R1").Select
Selection.Columns.ColumnWidth = 20
Range("S1").Select
Selection.Columns.ColumnWidth = 20
Sheets.Add After:=Sheets(Sheets.Count)
Sheets("Hoja3").Select
Sheets("Hoja3").Name = "UP_SELLING"
Windows("PLANTILLA.xlsm").Activate
Sheets("UP SELLING").Select
Set RangoDatos = Sheets("UP SELLING").UsedRange
RangoDatos.AutoFilter Field:=1, Criteria1:=valor
ultima = Sheets("UP SELLING").Range("A" & Rows.Count).End(xlUp).Row
Sheets("UP SELLING").Range("A1:Y" & ultima).Copy
Windows(ss).Activate
Sheets("UP_SELLING").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("T1").Select
Selection.Columns.ColumnWidth = 20
Range("W1").Select
Selection.Columns.ColumnWidth = 20
Range("X1").Select
Selection.Columns.ColumnWidth = 20
Sheets.Add After:=Sheets(Sheets.Count)
Sheets("Hoja4").Select
Sheets("Hoja4").Name = "RECONOCIMIENTO_LOGISTICO"
Windows("PLANTILLA.xlsm").Activate
Sheets("RECONOCIMIENTO LOGISTICO").Select
Set RangoDatos = Sheets("RECONOCIMIENTO LOGISTICO").UsedRange
RangoDatos.AutoFilter Field:=1, Criteria1:=valor
ultima = Sheets("RECONOCIMIENTO LOGISTICO").Range("A" & Rows.Count).End(xlUp).Row
Sheets("RECONOCIMIENTO LOGISTICO").Range("A1:Y" & ultima).Copy
Windows(ss).Activate
Sheets("RECONOCIMIENTO_LOGISTICO").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("N1").Select
Selection.Columns.ColumnWidth = 20
Range("O1").Select
Selection.Columns.ColumnWidth = 20
Range("P1").Select
Selection.Columns.ColumnWidth = 20
Range("S1").Select
Selection.Columns.ColumnWidth = 30
Range("T1").Select
Selection.Columns.ColumnWidth = 30
Range("U1").Select
Selection.Columns.ColumnWidth = 20
Sheets.Add After:=Sheets(Sheets.Count)
Sheets("Hoja5").Select
Sheets("Hoja5").Name = "CPS_"
Windows("PLANTILLA.xlsm").Activate
Sheets("CPS").Select
Set RangoDatos = Sheets("CPS").UsedRange
RangoDatos.AutoFilter Field:=1, Criteria1:=valor
ultima = Sheets("CPS").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CPS").Range("A1:Y" & ultima).Copy
Windows(ss).Activate
Sheets("CPS_").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("K1").Select
Selection.Columns.ColumnWidth = 20
Range("L1").Select
Selection.Columns.ColumnWidth = 20
Range("M1").Select
Selection.Columns.ColumnWidth = 20
Range("Q1").Select
Selection.Columns.ColumnWidth = 20
Range("P1").Select
Selection.Columns.ColumnWidth = 20
Sheets.Add After:=Sheets(Sheets.Count)
Sheets("Hoja6").Select
Sheets("Hoja6").Name = "CLARO_UP"
Windows("PLANTILLA.xlsm").Activate
Sheets("CLARO UP").Select
Set RangoDatos = Sheets("CLARO UP").UsedRange
RangoDatos.AutoFilter Field:=1, Criteria1:=valor
ultima = Sheets("CLARO UP").Range("A" & Rows.Count).End(xlUp).Row
Sheets("CLARO UP").Range("A1:Y" & ultima).Copy
Windows(ss).Activate
Sheets("CLARO_UP").Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Range("I1").Select
Selection.Columns.ColumnWidth = 20
Sheets("Hoja7").Delete
Sheets("CORRESPONSALIA_").Select
ActiveWorkbook.Save
ActiveWorkbook.Close
Windows("PLANTILLA.xlsm").Activate
Sheets("maestra").Select
Next i
MsgBox "DETALLES REALIZADOS CON ÉXITO", vbInformation
MsgBox "Por favor espere DOS minutos mientras se enfría el procesador, de lo contrario la memoria se desboradará", vbInformation
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub auto_open()
UserForm1.Show
End Sub
Sub uomo()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("maestra").Select
Range("A1").Select
river = Selection.End(xlDown).Row
Range("A2:G" & river).ClearContents
Range("A1").Select
Sheets("RESIDUAL").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A2").Select
river = Selection.End(xlDown).Row
Range("A2:N" & river).ClearContents
Range("A1").Select
Sheets("CORRESPONSALIA").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A2").Select
river = Selection.End(xlDown).Row
Range("A2:S" & river).ClearContents
Range("A1").Select
Sheets("UP SELLING").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A2").Select
river = Selection.End(xlDown).Row
Range("A2:Y" & river).ClearContents
Range("A1").Select
Sheets("RECONOCIMIENTO LOGISTICO").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A2").Select
river = Selection.End(xlDown).Row
Range("A2:X" & river).ClearContents
Range("A1").Select
Sheets("CPS").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A2").Select
river = Selection.End(xlDown).Row
Range("A2:R" & river).ClearContents
Range("A1").Select
Sheets("CLARO UP").Select
Range("A1").Select
ActiveSheet.ShowAllData
Range("A2").Select
river = Selection.End(xlDown).Row
Range("A2:M" & river).ClearContents
Range("A1").Select
Sheets("maestra").Select
Range("A1").Select
MsgBox "Plantilla de liquidación borrada con éxito, el ejecutable de excel se cerrará", vbInformation
Kill "D:\ARCHIVOS_AGENTES\ARCHIVOS\*.xlsx"
ActiveWorkbook.Close SaveChanges:=True
Application.Quit
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub marado()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim archivos
Dim goku As Excel.Workbook
Dim ws As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim nomlib As String
Set goku = Workbooks.Open("D:\ARCHIVOS_AGENTES\correos\CORREOS")
xiaomi = "CORREOS"
ss = xiaomi & ".xlsx"
Windows(ss).Activate
Windows("PLANTILLA.xlsm").Activate
Sheets("maestra").Select
Range("A1").Select
frank = Sheets("maestra").Range("A" & Rows.Count).End(xlUp).Row
Range("E2").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],[CORREOS.xlsx]Hoja1!C1:C11,11,FALSE)"
ActiveCell.Copy
Range("E3:E" & frank).Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("E:E").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Range("E1").Select
Windows(ss).Activate
ActiveWorkbook.Close
Windows("PLANTILLA.xlsm").Activate
MsgBox "Proceso Final Final completado puede continuar con el siguiente paso", vbInformation
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub gallardo()
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Dim FileQUITA As String
Dim napoli As CDO.Message
Windows("PLANTILLA.xlsm").Activate
Sheets("maestra").Select
tesl = Range("H1").Value
spax = Range("H2").Value
Range("a1").Select
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
FileQUITA = "D:\ARCHIVOS_AGENTES\ARCHIVOS" & "\" & Range("C" & i + 1) & ".xlsx"
para = Range("E" & i + 1)
asunto = Range("C" & i + 1)
lineaB = "En el archivo adjunto podrá encontrar el detalle de las liquidaciones de Residual, Up selling, Corresponsalía, Reconocimiento logístico, CPS y Claro UP."
lineaC = "<p></p><p>Cordialmente, Gerencia de comisiones.</p>"
With napoli
.Subject = asunto
.From = tesl
.To = para
.HTMLBody = lineaB & lineaC
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