Attribute VB_Name = "Module1"
Sub Sensitivity()

Dim live_case As Integer
Dim startrow As Integer
Dim startcol As Integer
Dim live_rows As Integer
Dim live_cols As Integer
Dim j As Integer
Dim sens_no As Integer
Dim senslist As Range

live_case = Range("Live_case").Value
startrow = Range("Sens_startrow").Value
startcol = Range("Sens_startcol").Value
live_rows = Range("Sens_live").Rows.Count
live_cols = Range("Sens_live").Columns.Count

sens_no = 0
Set senslist = Range("Sens_list")
For j = 1 To senslist.Count
    If senslist(j).Value <> 0 Then
        sens_no = sens_no + 1
    End If
Next j

For j = 1 To sens_no
    
    Range("Live_case").Value = j
    Call Sens_platformconsol
    Calculate
    
    With Sheets("Sensis")
        Range(.Cells(startrow, startcol), .Cells(startrow + live_rows - 1, startcol + live_cols - 1)) = Range("Sens_live").Value
    End With
    
    startrow = startrow + live_rows
    
Next j

Range("Live_case").Value = live_case
Call Sens_platformconsol

End Sub
Sub Sens_platformconsol()

Application.ScreenUpdating = False
    
    'Store starting sizing mode and switch it off
    sizing_switch = Range("Sizing_Active")
    If Range("Sens_override") = True Then
        Range("Sizing_Active") = 0
    End If
    
    'Store starting calculation mode and set to manual calculations
    calcmode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    'Start macro timer
    'Application.DisplayStatusBar = True
    'strStartTime = Now
    
    Dim projectList As Range
    Dim liveProject As Range
    Dim SPV_ConsolidatedCashflows As Range
    Dim rangeRowsCF As Integer
    Dim rangeColumnsCF As Integer
    Dim numberOfProjects As Integer
    Dim currentProjectNumber As Integer
    Dim i As Integer
    Dim test As Integer
    Dim startingrow As Integer
    Dim startingcolumn As Integer
    Set projectList = Range("ProjectList")
    Set projectactiveflag = Range("ProjectconsolidateFlag")
    Set liveProject = Range("Project_View")
    Set SPV_ConsolidatedCashflows = Range("SPV_ConsolidatedCashflows")
    
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
        Range(.Cells(startingrow, startingcolumn), .Cells(startingrow + 50000, startingcolumn + 500)).Clear
    End With
   
    
       For i = 1 To projectList.Count And projectactiveflag.Count
           If projectList(i) <> 0 And projectactiveflag(i) = True Then
                currentProjectNumber = currentProjectNumber + 1
                'dtDuration = DateDiff("n", CDate(strStartTime), CDate(strCompTime))
                
                liveProject.Value = projectList(i)
                'Application.StatusBar = "Updating project " & currentProjectNumber & " of " & numberOfProjects & " - " & liveProject.Value & ". Time elapsed " & dtDuration & "mins."
                
                'Solve circularities for individual project
                'Call Debt_sizing
                
                Application.Calculation = xlCalculationManual
                Calculate
                                

                
                
               
                'Paste live results into into output sheet
                With Sheets("quarterly output")
                    Range(.Cells(startingrow, startingcolumn), .Cells(startingrow + rangeRowsCF - 1, startingcolumn + rangeColumnsCF - 1)) = SPV_ConsolidatedCashflows.Value
                End With
                startingrow = startingrow + rangeRowsCF
                
                Application.ScreenUpdating = True
                strCompTime = Now
           End If
           
       Next i
   
    'MsgBox "Total time taken to complete consolidation " & dtDuration & "mins.", vbOKOnly, "Consolidation completed"
    
    'Application.StatusBar = ""
    Application.ScreenUpdating = True
    'Reset to starting sizing mode and calculation mode
    Range("Sizing_Active") = sizing_switch
    Application.Calculation = calcmode
    
End Sub

