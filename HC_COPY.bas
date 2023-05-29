Attribute VB_Name = "HC_COPY"
Sub HC_copy()

    Dim original_workbook As Workbook
    Dim HC_Workbook As Workbook
    
    Dim HC_WorkbookPath As String
    Dim HCName As String
    
    Set original_workbook = ThisWorkbook
    
    ' NAME OF THE HC WORKBOOK
    HCName = "HC_" & VBA.Format(Now, "dd_mm_yy__hh_mm") & "_" & ThisWorkbook.Name
    
    'COMPLETE PATH OF HC WORKBOOK
    HC_WorkbookPath = ThisWorkbook.Path & "\" & HCName
    
    ' CREATING THE COPY TO EDIT
    ThisWorkbook.SaveCopyAs HC_WorkbookPath
    
    ' OPENING THE HC COPY OF WORKBOOK
    Set HC_Workbook = Workbooks.Open(HC_WorkbookPath)
    
    ' EDITING THE HC WORKBOOK -> CREATING THE HC COPY
    Sheets.Select
    Cells.Copy
    Cells.PasteSpecial xlPasteValues
    
    Range("A1").Select
    Sheets(1).Select
    
    ' SAVING AND CLOSING HC WORKBOOK
    HC_Workbook.Save
    HC_Workbook.Close
    
    
    ' RELEASE OBJECT REFERENCES
    Set original_workbook = Nothing
    Set HC_Workbook = Nothing
    
End Sub
