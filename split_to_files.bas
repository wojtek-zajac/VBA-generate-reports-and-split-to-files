Sub PARTNERS_split_to_files()


'/////////////////////////////////////////////////////////////////////////////////////////////////////
'IMPORTANT!
'This macro won't work stored in "Personal Macro Workbook". Please put it to This Workbook’s Module.
'/////////////////////////////////////////////////////////////////////////////////////////////////////


Application.ScreenUpdating = False


Dim wbThis As Workbook
Dim wbNew As Workbook
Dim ws As Worksheet
Dim strFilename As String


    Set wbThis = ThisWorkbook
    For Each ws In wbThis.Worksheets
        strFilename = ws.Name & "_" & "11-2016" & "_Report.xlsx"
        
 '/////////////////////////////////////////////////////////////////////////////////////////////////////
 'IMPORTANT
 'Please change the date (i.e. for 09-2016) accordingly.
 '/////////////////////////////////////////////////////////////////////////////////////////////////////
 
        ws.Copy
        Set wbNew = ActiveWorkbook
        wbNew.SaveAs strFilename
        wbNew.Close
    Next ws
    
Application.ScreenUpdating = True


End Sub
