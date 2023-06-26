Attribute VB_Name = "Módulo1"

' SCRIPT TO CHOOSE THE DELIMITER AND THE RANGE

Sub export_to_file()

    ' DECLARATION OF VARIABLES
    Dim RangeToExport As Range
    Dim Delimiter As String
    Dim FileName As String
    Dim ExpRow As Range
    Dim ExpCell As Range
    Dim MyValue As Variant
    
    'NAME OF THE NEW FILE
    FileName = ThisWorkbook.Path & "\NewFile.txt"
    
    'OPENING THE FILE
    Open FileName For Output As #1
    
    'INPUT OF THE RANGE AND DELIMITER
    On Error GoTo End_Sub
    Set RangeToExport = Application.InputBox("Please select a cell in the range of values:", "Select Range", Type:=8)
    
    ' INPUT THE DELIMITER
    Delimiter = InputBox("Select the Delimiter:", "Delimiter Selection")
    
    'ITERATE THROUGH THE ROWS TO LOOP IN THE CELLS
    For Each ExpRow In RangeToExport.CurrentRegion.Rows
        For Each ExpCell In ExpRow.Cells
            MyValue = MyValue & ExpCell.Value & Delimiter
            Next ExpCell
            
            'WRITE IN THE FILE
            Print #1, MyValue
            
        MyValue = ""
        Next ExpRow
    
        'CLOSE THE FILE
        Close #1
    
End_Sub:
End Sub
