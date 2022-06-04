Sub Click_ResetTool()

Dim sht As Worksheet
Dim Log As Excel.Worksheet, Questions As Excel.Worksheet
Dim Output_Data_Start_Row As Long: Output_Data_Start_Row = 2
Dim Output_Data_Start_Col As Long: Output_Data_Start_Col = 2
Dim Max_Clear_Row_Count As Long: Max_Clear_Row_Count = 1000000

Application.DisplayAlerts = False

Set Log = ThisWorkbook.Sheets("Log")
Set Questions = ThisWorkbook.Sheets("Questions")

Log_Header_Row = Log.Cells.Find("_Import Log_", , , xlWhole).Offset(1, 0).Row
Log_DataStart_Row = Log.Cells.Find("_Import Log_", , , xlWhole).Offset(2, 0).Row
Log_DataStart_Col = Log.Cells.Find("_Import Log_", , , xlWhole).Column
Log_DataEnd_Col = Log.Cells(Log_Header_Row, Log_DataStart_Col).End(xlToRight).Column

If MsgBox("Are you sure you want to reset the tool?" & vbCrLf, vbYesNo) = vbNo Then
    Exit Sub
End If
Log.Range(Log.Cells(Log_DataStart_Row, Log_DataStart_Col), Log.Cells(Max_Clear_Row_Count, Log_DataEnd_Col)).ClearContents
' clear questions sheet
Set Questions = ThisWorkbook.Worksheets("Questions")
        Output_Data_End_Col = Questions.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlToRight).Column
        Questions.Range(Questions.Cells(Output_Data_Start_Row + 1, Output_Data_Start_Col), Questions.Cells(Max_Clear_Row_Count, Output_Data_End_Col)).ClearContents

For Each sht In ThisWorkbook.Worksheets
    If Not (sht.Name = "Log" Or sht.Name = "Questions" Or sht.Name = "Table Template Sheet" Or sht.Name = "1. Example Format") Then sht.Delete
Next sht

Application.DisplayAlerts = True

ThisWorkbook.Worksheets("Table Template Sheet").Visible = xlSheetVisible
ThisWorkbook.Worksheets("1. Example Format").Visible = xlSheetVisible

Log.Activate: Log.Cells(1, 1).Select

MsgBox "All data in the sheets ""Log"" and ""Questions"" cleared. All the extra table sheets have been deleted."

End Sub

Sub Click_CreateTableSheets()
    Dim fd As FileDialog, xlApp As Excel.Application, xlWkb As Excel.Workbook, Wks As Excel.Worksheet
    Dim Log As Excel.Worksheet, Output As Excel.Worksheet, RowCounters As Excel.Worksheet
    Dim StartTime As Double
    Dim Day, File_Path, TableTemplateSheetName As String
    Dim Output_Data_Start_Row As Long: Output_Data_Start_Row = 2
    Dim Output_Data_Start_Col As Long: Output_Data_Start_Col = 2
    Dim Max_Clear_Row_Count As Long: Max_Clear_Row_Count = 1000000
    Err_Count = 0
    TableTemplateSheetName = "Table Template Sheet"
    
    ' Setting Worksheets
    Application.ScreenUpdating = False
    Set Log = ThisWorkbook.Sheets("Log")
    Set ToolWkb = ThisWorkbook
    
    ' Finds table dimensions for the Log table and the Question and Table Config table
    Config_Table_Header_Row = Log.Cells.Find("_Question And Table Config_", , , xlWhole).Offset(1, 0).Row
    Config_Table_DataStart_Row = Log.Cells.Find("_Question And Table Config_", , , xlWhole).Offset(2, 0).Row
    Config_Table_DataStart_Col = Log.Cells.Find("_Question And Table Config_", , , xlWhole).Column
    Config_Table_DataEnd_Col = Log.Cells(Config_Table_Header_Row, Config_Table_DataStart_Col).End(xlToRight).Column
    Config_Table_DataEnd_Row = Log.Cells(Config_Table_Header_Row, Config_Table_DataStart_Col).End(xlDown).Row
    ' for both questions and tables
    Config_Table_Question_Or_Table_Col = Config_Table_DataStart_Col
    Config_Table_Sheet_Name_Col = Config_Table_DataStart_Col + 1
    ' for only tables
    Config_Table_Sheet_Top_Left_Col = Config_Table_DataStart_Col + 10
    Config_Table_First_Row_Offset_Col = Config_Table_DataStart_Col + 11
    Config_Table_Sheet_Output_Col = Config_Table_DataStart_Col + 12
    
    '   Select files to consolidate.
     MsgBox "Please select your template file for table sheet creation."
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.InitialView = msoFileDialogViewList
    fd.AllowMultiSelect = False
    FileChosen = fd.Show
        If FileChosen = 0 Then
            MsgBox "No files selected...Try again", vbInformation
            Exit Sub
        End If
    StartTime = Timer
    Set xlApp = GetExcelObject
    xlApp.DisplayAlerts = False
'    xlApp.Visible = True
                          
    ' Cycles through selected files and starts consolidation
    For fi = 1 To fd.SelectedItems.Count
        File_Path = fd.SelectedItems(fi)
        Set xlWkb = xlApp.Workbooks.Open(File_Path)
        Dim Supplier_Name As String: Supplier_Name = ""
        Supplier_Name = Split(xlWkb.Name, "_")(0)
        Wkb_Name = xlWkb.Name
        For Ind = Config_Table_DataStart_Row To Config_Table_DataEnd_Row
            If Log.Cells(Ind, Config_Table_Question_Or_Table_Col) = "Table" Then
                ' table capture
                On Error GoTo Error_Handling
                Set Wks = xlWkb.Worksheets(CStr(Log.Cells(Ind, Config_Table_Sheet_Name_Col)))
                ToolWkb.Worksheets(TableTemplateSheetName).Copy After:=ToolWkb.Sheets(ToolWkb.Worksheets.Count)
                Set Output = ActiveSheet
                Output.Name = Log.Cells(Ind, Config_Table_Sheet_Output_Col)
                Output_PasteStart_Col = Output.Cells.Find("Question Text", , , xlWhole).Offset(0, 1).Column
                
                Set Wks = xlWkb.Worksheets(CStr(Log.Cells(Ind, Config_Table_Sheet_Name_Col)))
                Top_Left = Log.Cells(Ind, Config_Table_Sheet_Top_Left_Col)
                
                Wks_Top_Left_Row = Wks.Range(Top_Left).Row
                Wks_Top_Left_Col = Wks.Range(Top_Left).Column
                Wks_DataEnd_Col = Wks.Range(Top_Left).End(xlToRight).Column
                Wks_Col_Count = Wks_DataEnd_Col - Wks_Top_Left_Col
                ' copy and paste
                Output.Range(Output.Cells(Output_Data_Start_Row, Output_PasteStart_Col), Output.Cells(Output_Data_Start_Row, Output_PasteStart_Col + Wks_Col_Count)).Value2 = _
                Wks.Range(Wks.Cells(Wks_Top_Left_Row, Wks_Top_Left_Col), Wks.Cells(Wks_Top_Left_Row, Wks_DataEnd_Col)).Value2
                
                Output_DataEnd_Col = Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlToRight).Column
                
                Start_Col_Letter = Split(Cells(1, Output_Data_Start_Col).Address, "$")(1)
                Output.Columns(Start_Col_Letter & ":" & Start_Col_Letter).Copy
                PasteStart_Col_Letter = Split(Cells(1, Output_PasteStart_Col).Address, "$")(1)
                DataEnd_Col_Letter = Split(Cells(1, Output_DataEnd_Col).Address, "$")(1)
                Output.Columns(PasteStart_Col_Letter & ":" & DataEnd_Col_Letter).PasteSpecial Paste:=xlPasteFormats
                Output.Columns(PasteStart_Col_Letter & ":" & DataEnd_Col_Letter).AutoFit
                Application.CutCopyMode = False
                Output.Range(Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col), Output.Cells(Output_Data_Start_Row, Output_DataEnd_Col)).AutoFilter
            End If
        Next Ind
        Debug.Print Supplier_Name & ": " & "....Complete"
        xlWkb.Close
        Set xlWkb = Nothing
        ' Error handling
Error_Handling:
     If Supplier_Name = Wkb_Name Then
        Error_Comment = "Cannot find a supplier name in the file [" & Wkb_Name & "]. Please add text before the first underscore and re-upload"
        MsgBox Error_Comment
        xlWkb.Close
        Set xlWkb = Nothing
        Err_Count = Err_Count + 1
     ElseIf Err.Number = 9 Then
        Error_Comment = "Cannot find a sheet called " & CStr(Log.Cells(Ind, Config_Table_Sheet_Name_Col)) & " in this workbook, please adjust and re-upload"
        MsgBox Error_Comment
        Err.Clear
        Err_Count = Err_Count + 1
     ElseIf Err.Number <> 0 Then
        Error_Comment = Err.Description
        MsgBox Error_Comment
        Err.Clear
        Err_Count = Err_Count + 1
     End If
     
    File_ID = File_ID + 1
    Next fi
    xlApp.DisplayAlerts = True
    xlApp.Quit
    Set xlApp = Nothing
    ToolWkb.Worksheets(TableTemplateSheetName).Visible = xlSheetVeryHidden
    ToolWkb.Worksheets("1. Example Format").Visible = xlSheetVeryHidden
    Log.Activate: Log.Cells(1, 1).Select
    ' Final message box showing time taken and error files
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    Application.ScreenUpdating = True
    If Err_Count > 0 Then
        MsgBox "Errors during sheet creation detected. Please review and try again."
    Else
        MsgBox "Sheet Creation Complete. It took: " & MinutesElapsed & " hh:mm:ss."
    End If
        
End Sub


Public Function GetExcelObject() As Object
    On Error Resume Next
    Dim xlo As Object
    'Set xlo = GetObject("Excel.Application")
    If xlo Is Nothing Then
        Set xlo = CreateObject("Excel.Application")
    End If
    Set GetExcelObject = xlo
End Function
