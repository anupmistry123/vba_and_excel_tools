Sub Click_Consolidate_Bid_Files()

Call ConsolidateTemplate.Consolidate(False)

End Sub

Sub Click_Clear_All_Data_Button()

Call ConsolidateTemplate.Consolidate(True)

End Sub

Sub Consolidate(Optional Click_Clear_All_Data As Boolean = False)
    Dim fd As FileDialog, xlApp As Excel.Application, xlWkb As Excel.Workbook, Wks As Excel.Worksheet
    Dim Log As Excel.Worksheet, Output As Excel.Worksheet, RowCounters As Excel.Worksheet
    Dim StartTime As Double
    Dim Day, File_Path As String
    Dim Output_Data_Start_Row As Long: Output_Data_Start_Row = 2
    Dim Output_Data_Start_Col As Long: Output_Data_Start_Col = 2
    Dim Max_Clear_Row_Count As Long: Max_Clear_Row_Count = 1000000
    Dim QuestionData(1 To 9) As Variant
    Err_Count = 0
    
    ' Setting Worksheets
    Application.ScreenUpdating = False
    Set Log = ThisWorkbook.Sheets("Log")
    
    ' Finds table dimensions for the Log table and the Question and Table Config table
    Log_Header_Row = Log.Cells.Find("_Import Log_", , , xlWhole).Offset(1, 0).Row
    Log_DataStart_Row = Log.Cells.Find("_Import Log_", , , xlWhole).Offset(2, 0).Row
    Log_DataStart_Col = Log.Cells.Find("_Import Log_", , , xlWhole).Column
    Log_DataEnd_Col = Log.Cells(Log_Header_Row, Log_DataStart_Col).End(xlToRight).Column
    
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
    Config_Table_Bottom_Row_Col = Config_Table_DataStart_Col + 13
    
    ' Clears Data
    If Click_Clear_All_Data Then
        If MsgBox("Are you sure you want to clear all data?" & vbCrLf, vbYesNo) = vbNo Then
            Exit Sub
        End If
        Log.Range(Log.Cells(Log_DataStart_Row, Log_DataStart_Col), Log.Cells(Max_Clear_Row_Count, Log_DataEnd_Col)).ClearContents
        ' clear questions sheet
        Set Output = ThisWorkbook.Worksheets("Questions")
                Output_Data_End_Col = Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlToRight).Column
                Output.Range(Output.Cells(Output_Data_Start_Row + 1, Output_Data_Start_Col), Output.Cells(Max_Clear_Row_Count, Output_Data_End_Col)).ClearContents
        ' clear table sheets
        For Ind = Config_Table_DataStart_Row To Config_Table_DataEnd_Row
            If Log.Cells(Ind, Config_Table_Question_Or_Table_Col).Value = "Table" Then
                Set Output = ThisWorkbook.Worksheets(CStr(Log.Cells(Ind, Config_Table_Sheet_Output_Col)))
                Output_Data_End_Col = Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlToRight).Column
                Output.Range(Output.Cells(Output_Data_Start_Row + 1, Output_Data_Start_Col), Output.Cells(Max_Clear_Row_Count, Output_Data_End_Col)).ClearContents
            End If
        Next Ind
        MsgBox "All data in all sheets cleared"
        Exit Sub
    End If
    
    '   Select files to consolidate.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.InitialView = msoFileDialogViewList
    fd.AllowMultiSelect = True
    FileChosen = fd.Show
        If FileChosen = 0 Then
            MsgBox "No files selected...Try again", vbInformation
            Exit Sub
        End If
    StartTime = Timer
    Set xlApp = GetExcelObject
    xlApp.DisplayAlerts = False
'    xlApp.Visible = True
    
    ' Set File_ID
    File_ID = WorksheetFunction.Max(Log.Columns(Log_DataStart_Col)) + 1
                          
    ' Cycles through selected files and starts consolidation
    For fi = 1 To fd.SelectedItems.Count
        ' Add a row to the Import Log
        File_Path = fd.SelectedItems(fi)
        Set xlWkb = xlApp.Workbooks.Open(File_Path)
        Timestamp = Format(Now(), "DD/MM/YYYY hh:mm:ss")
        Import_User = Environ("username")
        Dim Supplier_Name As String: Supplier_Name = ""
        Supplier_Name = Split(xlWkb.Name, "_")(0)
        Wkb_Name = xlWkb.Name
        If Supplier_Name = Wkb_Name Then GoTo Error_Handling
        
        For Ind = Config_Table_DataStart_Row To Config_Table_DataEnd_Row
            Set Wks = xlWkb.Worksheets(CStr(Log.Cells(Ind, Config_Table_Sheet_Name_Col)))
            ArrayNum = 0
                QuestionData(LBound(QuestionData)) = IIf(Log.Cells(Ind, Config_Table_DataStart_Col + LBound(QuestionData)) = "", "", Log.Cells(Ind, Config_Table_DataStart_Col + LBound(QuestionData)))
                For ArrayNum = LBound(QuestionData) + 1 To UBound(QuestionData)
                    If Log.Cells(Ind, Config_Table_DataStart_Col + ArrayNum) = "" Then
                        QuestionData(ArrayNum) = ""
                    Else
                        QuestionData(ArrayNum) = Wks.Range(CStr(Log.Cells(Ind, Config_Table_DataStart_Col + ArrayNum)))
                    End If
                Next ArrayNum
            If Log.Cells(Ind, Config_Table_Question_Or_Table_Col) = "Question" Then
                ' question capture
                Set Output = ThisWorkbook.Worksheets("Questions")
                Output_DataEnd_Row = IIf(Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlDown).Row = Output.Rows.Count, 2, Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlDown).Row)
                Output_PasteStart_Col = Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).Offset(0, 2).Column
                Output_DataEnd_Col = Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlToRight).Column
                
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_PasteStart_Col), Output.Cells(Output_DataEnd_Row + 1, Output_DataEnd_Col)).Value2 = QuestionData()
                ' Erase QuestionData
                
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col), Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col)) = File_ID
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col + 1), Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col + 1)) = Supplier_Name
            Else
                ' table capture
                Set Output = ThisWorkbook.Worksheets(CStr(Log.Cells(Ind, Config_Table_Sheet_Output_Col)))
                Output_DataEnd_Row = IIf(Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlDown).Row = Output.Rows.Count, 2, Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlDown).Row)
                Output_PasteStart_Col = Output.Cells.Find("Question Text", , , xlWhole).Offset(0, 1).Column
                Output_DataEnd_Col = Output.Cells(Output_Data_Start_Row, Output_Data_Start_Col).End(xlToRight).Column
                On Error GoTo Error_Handling
                Set Wks = xlWkb.Worksheets(CStr(Log.Cells(Ind, Config_Table_Sheet_Name_Col)))
                Top_Left = Log.Cells(Ind, Config_Table_Sheet_Top_Left_Col)
                First_Row_Offset = Log.Cells(Ind, Config_Table_First_Row_Offset_Col)
                Wks_Top_Left_Row = Wks.Range(Top_Left).Row
                Wks_Top_Left_Col = Wks.Range(Top_Left).Column
                Wks_DataStart_Row = Wks.Range(Top_Left).Row + First_Row_Offset
                If Log.Cells(Ind, Config_Table_Bottom_Row_Col) = "" Or Log.Cells(Ind, Config_Table_Bottom_Row_Col) = 0 Then
                    Wks_DataEnd_Row = IIf(Wks.Cells(Wks_DataStart_Row, Wks_Top_Left_Col).End(xlDown).Row <> Wks.Rows.Count, Wks.Cells(Wks_DataStart_Row, Wks_Top_Left_Col).End(xlDown).Row, 2)
                Else
                    Wks_DataEnd_Row = Log.Cells(Ind, Config_Table_Bottom_Row_Col)
                End If
                Wks_DataEnd_Col = Wks.Range(Top_Left).End(xlToRight).Column
                Wks_DataRowCount = (Wks_DataEnd_Row - Wks_DataStart_Row) + 1
                ' copy and paste
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_PasteStart_Col), Output.Cells(Output_DataEnd_Row + Wks_DataRowCount, Output_DataEnd_Col)).Value2 = _
                Wks.Range(Wks.Cells(Wks_DataStart_Row, Wks_Top_Left_Col), Wks.Cells(Wks_DataEnd_Row, Wks_DataEnd_Col)).Value2
                
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col), Output.Cells(Output_DataEnd_Row + Wks_DataRowCount, Output_Data_Start_Col)) = File_ID
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col + 1), Output.Cells(Output_DataEnd_Row + Wks_DataRowCount, Output_Data_Start_Col + 1)) = Supplier_Name
                Output.Range(Output.Cells(Output_DataEnd_Row + 1, Output_Data_Start_Col + 2), Output.Cells(Output_DataEnd_Row + Wks_DataRowCount, Output_PasteStart_Col - 1)) = QuestionData()
                
                
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
        Log.Cells(Log_Header_Row + File_ID, 7) = Error_Comment
        Err_Count = Err_Count + 1
     ElseIf Err.Number = 9 Then
        Error_Comment = "Cannot find a sheet called " & CStr(Log.Cells(Ind, Config_Table_Sheet_Name_Col)) & " in this workbook, please adjust and re-upload"
        MsgBox Error_Comment
        Log.Cells(Log_Header_Row + File_ID, 7) = Error_Comment
        Err.Clear
        Err_Count = Err_Count + 1
     ElseIf Err.Number = 91 Then
        Error_Comment = "Cannot find one of the header columns in this workbook, please adjust and re-upload"
        MsgBox Error_Comment
        Log.Cells(Log_Header_Row + File_ID, 7) = Error_Comment
        Err.Clear
        Err_Count = Err_Count + 1
     ElseIf Err.Number <> 0 Then
        Error_Comment = Err.Description
        MsgBox Error_Comment
        Log.Cells(Log_Header_Row + File_ID, 7) = Error_Comment
        Err.Clear
        Err_Count = Err_Count + 1
     End If
     
    Log.Cells(Log_Header_Row + File_ID, 2) = File_ID
    Log.Cells(Log_Header_Row + File_ID, 3) = File_Path
    Log.Cells(Log_Header_Row + File_ID, 4) = Timestamp
    Log.Cells(Log_Header_Row + File_ID, 5) = Import_User
    Log.Cells(Log_Header_Row + File_ID, 6) = Supplier_Name
     
    File_ID = File_ID + 1
    Next fi
    xlApp.DisplayAlerts = True
    xlApp.Quit
    Set xlApp = Nothing
    ' Final message box showing time taken and error files
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    Application.ScreenUpdating = True
    MsgBox "Consolidation Complete. Consolidated " & fd.SelectedItems.Count & " files in: " & MinutesElapsed & " hh:mm:ss." & vbCrLf & _
           Err_Count & " files had errors. Please review the Log table and re-upload if required."
        
End Sub