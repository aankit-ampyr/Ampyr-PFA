Attribute VB_Name = "Consolidation"
Public Sub PlatformConsolidation()
        
    Application.ScreenUpdating = False
    
    'Store starting calculation mode and set to manual calculations
    calcmode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    'Start macro timer
    Application.DisplayStatusBar = True
    strStartTime = Now
    
    Dim projectList As Range
    Dim liveProject As Range
    Dim SPV_ConsolidatedCashflows As Range
    Dim rangeRowsCF As Integer
    Dim rangeColumnsCF As Integer
    Dim numberOfProjects As Integer
    Dim currentProjectNumber As Integer
    Dim livecase As Integer
    Dim i As Integer
    Dim test As Integer
    Dim startingrow As Integer
    Dim startingcolumn As Integer
    Set projectList = Range("ProjectList")
    'Set projectactiveflag = Range("ProjectconsolidateFlag")
    Set projectactiveflag = Range("ProjectActiveFlag")
    Set liveProject = Range("Project_View")
    Set SPV_ConsolidatedCashflows = Range("SPV_ConsolidatedCashflows")
    
    livecase = Range("Live_case").Value
    Range("Live_case").Value = 1
    'Count number of projects filled in
    For i = 1 To projectList.Count
        If projectList(i).Value <> 0 Then
                numberOfProjects = numberOfProjects + 1
        End If
    Next i

    'Store output start column and row number
    'outputCellStart = Split(Range("outputCellStart").Value, "$") '
    'outputColumn = outputCellStart(1)
    'outputRow = outputCellStart(2)
    
    'Store output start column and row number for sale
    'outputSaleCellStart = Split(Range("outputSaleCellStart").Value, "$")
    'outputSaleColumn = outputSaleCellStart(1)
    'outputSaleRow = outputSaleCellStart(2)
    
    'Clear the output tab


    
    startingrow = Range("PasteStartRow").Value
    startingcolumn = Range("pastestartcolumn").Value
    rangeRowsCF = SPV_ConsolidatedCashflows.Rows.Count
    rangeColumnsCF = SPV_ConsolidatedCashflows.Columns.Count

    'Clear previous pasted outputs
    With Sheets("quarterly output")
        Range(.Cells(startingrow, startingcolumn), .Cells(startingrow + 90000, startingcolumn + 500)).Clear
    End With
   
    
       For i = 1 To projectList.Count + 1 And projectactiveflag.Count
           'If projectList(i) <> 0 And projectactiveflag(i) = True Then
          
           If projectList(i) <> 0 And projectactiveflag(i) = "True" Then
                currentProjectNumber = currentProjectNumber + 1
                dtDuration = DateDiff("n", CDate(strStartTime), CDate(strCompTime))
                
                liveProject.Value = projectList(i)
                Application.StatusBar = "Updating project " & currentProjectNumber & " of " & numberOfProjects & " - " & liveProject.Value & ". Time elapsed " & dtDuration & "mins."
                
                'Solve circularities for individual project
                Call Debt_sizing
                'Calculate
                Application.Calculation = xlCalculationManual
                                
            
                
               
                'Paste live results into into output sheet
                With Sheets("quarterly output")
                    Range(.Cells(startingrow, startingcolumn), .Cells(startingrow + rangeRowsCF - 1, startingcolumn + rangeColumnsCF - 1)) = SPV_ConsolidatedCashflows.Value
                End With
                startingrow = startingrow + rangeRowsCF
                
                Application.ScreenUpdating = True
                strCompTime = Now
                
           End If
           
       Next i
   
    MsgBox "Total time taken to complete consolidation " & dtDuration & "mins.", vbOKOnly, "Consolidation completed"
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    'Reset to starting calculation mode
    Calculate
    Application.Calculation = calcmode
    Range("Live_case").Value = livecase
    
Exit Sub

ErrorHandler:
    MsgBox "Exiting Macro due to Error: Check all projects are populated"
    Exit Sub
    
End Sub


Public Sub UpdateProject()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    
    Dim projectList As Range
    Dim liveProject As Range
    Set projectactiveflag = Range("ProjectActiveFLag")
    Set projectList = Range("ProjectList")
    Set liveProject = Range("Project_View")
    Set SPV_ConsolidatedCashflows = Range("SPV_ConsolidatedCashflows")
    Application.StatusBar = "Updating project " & liveProject & "...please wait"
    
    'Store output start column and row number
    outputColumn = Range("pastestartcolumn").Value
    outputRow = Range("PasteStartRow").Value
    rangeRowsCF = SPV_ConsolidatedCashflows.Rows.Count
    rangeColumnsCF = SPV_ConsolidatedCashflows.Columns.Count
    
    'Find address of project in the output tab
    For i = 1 To projectList.Count And projectactiveflag.Count
           If projectList(i) = liveProject And projectactiveflag(i) = "True" Then
            
            outputRow = rangeRowsCF * (i - 1) + outputRow
            'outputAddress = outputColumn & outputRow
            Exit For
        End If
    Next i
    
   Call Debt_sizing
    'Call Module7.Junior_sizing
    Application.Calculate
    Application.ScreenUpdating = False
    
    'SPV_ConsolidatedCashflows.Copy
    'Worksheets(Range("outputSheetName").Value).Range(outputAddress).PasteSpecial Paste:=xlPasteValues
    'Worksheets(Range("outputSheetName").Value).Range(outputAddress).PasteSpecial Paste:=xlPasteFormats
    With Sheets("quarterly output")
          Range(.Cells(outputRow, outputColumn), .Cells(outputRow + rangeRowsCF - 1, outputColumn + rangeColumnsCF - 1)) = SPV_ConsolidatedCashflows.Value
    End With
    Application.StatusBar = ""
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
End Sub

Public Sub RemoveProject()

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    
    Dim projectList As Range
    Dim liveProject As Range
    
    Set projectList = Range("ProjectList")
    Set liveProject = Range("Project_View")
    Set SPV_ConsolidatedCashflows = Range("SPV_ConsolidatedCashflows")
    Application.StatusBar = "Removing project " & liveProject
    
    'Store output start column and row number
    outputCellStart = Split(Range("outputCellStart").Value, "$")
    outputColumn = outputCellStart(1)
    outputRow = outputCellStart(2)
    
    'Find address of project in the output tab
    For i = 1 To projectList.Count
        If projectList(i) = liveProject Then
            rangeRowsCF = SPV_ConsolidatedCashflows.Rows.Count
            outputRow = rangeRowsCF * (i - 1) + 2 * (i - 1) + outputRow
            outputAddress = outputColumn & outputRow & ":" & "KQ" & outputRow + rangeRowsCF
            Exit For
        End If
    Next i
    
    Worksheets(Range("outputSheetName").Value).Range(outputAddress).Clear
    Application.StatusBar = ""
    Application.CutCopyMode = False
    Application.ScreenUpdating = True


End Sub '




