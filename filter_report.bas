Sub PARTNERS_REPORTS()


Application.ScreenUpdating = False
ActiveSheet.Name = "Consumption_Report"

Range("A:AD").Sort Key1:=Range("C1"), Header:=xlYes


[A:A, B:B, D:D, E:E, F:F, G:G, H:H, I:I, J:J, O:O, P:P, Q:Q, R:R, S:S, T:T, Z:Z, AC:AC, AD:AD].Delete


[1:1].Font.Bold = True


    Sheets.Add
    Sheets("Sheet1").Name = "Sonru"
    Sheets.Add
    Sheets("Sheet2").Name = "Workhoppers.com"
    Sheets.Add
    Sheets("Sheet3").Name = "VeteranCareer.org"
    Sheets.Add
    Sheets("Sheet4").Name = "iHireVeteran.com"
    Sheets.Add
    Sheets("Sheet5").Name = "USAJOBBOARD"
    Sheets.Add
    Sheets("Sheet6").Name = "ITJobX.com"
    Sheets.Add
    Sheets("Sheet7").Name = "Active Job Board"
    Sheets.Add
    Sheets("Sheet8").Name = "Totallyhired inc."
    Sheets.Add
    Sheets("Sheet9").Name = "SalesGravy"
    Sheets.Add
    Sheets("Sheet10").Name = "Recroup"
    Sheets.Add
    Sheets("Sheet11").Name = "Performance Assessment Network"
    Sheets.Add
    Sheets("Sheet12").Name = "PURE JOBS"
    Sheets.Add
    Sheets("Sheet13").Name = "LevoLeague"
    Sheets.Add
    Sheets("Sheet14").Name = "JustJobs.com"
    Sheets.Add
    Sheets("Sheet15").Name = "ITJobCafe"
    Sheets.Add
    Sheets("Sheet16").Name = "GlassDoorPro"
    Sheets.Add
    Sheets("Sheet17").Name = "Geebo"
    Sheets.Add
    Sheets("Sheet18").Name = "Good&Co"
    Sheets.Add
    Sheets("Sheet19").Name = "FashionUnited"
    Sheets.Add
    Sheets("Sheet20").Name = "Engineer Nexus LLC"
    Sheets.Add
    Sheets("Sheet21").Name = "DiversityJobs"
    Sheets.Add
    Sheets("Sheet22").Name = "Bio Careers"
    Sheets.Add
    Sheets("Sheet23").Name = "AccountantJobs.com"
    Sheets.Add
    Sheets("Sheet24").Name = "Chequed.com"
    Sheets.Add
    Sheets("Sheet25").Name = "Outmatch"
    
    
    
    
        
Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Sonru"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Sonru").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Workhoppers.com"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Workhoppers.com").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="VeteranCareer.org"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("VeteranCareer.org").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="iHireVeteran.com"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("iHireVeteran.com").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="USAJOBBOARD"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("USAJOBBOARD").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="ITJobX.com"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("ITJobX.com").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Active Job Board"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Active Job Board").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Totallyhired inc."
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Totallyhired inc.").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="SalesGravy"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("SalesGravy").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Recroup"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Recroup").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Performance Assessment Network"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Performance Assessment Network").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="PURE JOBS"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("PURE JOBS").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="LevoLeague"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("LevoLeague").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="JustJobs.com"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("JustJobs.com").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="ITJobCafe"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("ITJobCafe").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="GlassDoorPro"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("GlassDoorPro").Select
ActiveSheet.Paste
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Geebo"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Geebo").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True

Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Good&Co"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Good&Co").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True

Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="FashionUnited"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("FashionUnited").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Engineer Nexus LLC"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Engineer Nexus LLC").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="DiversityJobs"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("DiversityJobs").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True



Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Bio Careers"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Bio Careers").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="AccountantJobs.com"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("AccountantJobs.com").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Chequed.com"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Chequed.com").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Sheets("Consumption_Report").Select
Application.CutCopyMode = False
    Selection.AutoFilter
Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("A:L").AutoFilter Field:=2, Criteria1:="Outmatch"
    ActiveSheet.Range("A:L").AutoFilter Field:=10, Criteria1:="SUCCESS"
Range("A1:L" & Cells(Rows.Count, "A").End(xlUp).Row).Select
Selection.Copy
Sheets("Outmatch").Select
ActiveSheet.Paste
Application.CutCopyMode = False
Columns("A:L").EntireColumn.AutoFit
Application.DisplayAlerts = False
If LenB(ActiveSheet.Range("A2")) = 0 Then ActiveSheet.Delete
Application.DisplayAlerts = True


Application.DisplayAlerts = False
Sheets("Consumption_Report").Delete
Application.DisplayAlerts = True


Application.ScreenUpdating = True

MsgBox "YOLOOO!"

End Sub
