Attribute VB_Name = "DebtSizing"
Sub Debt_sizing()

Application.ScreenUpdating = False
calcmode = Application.Calculation
Application.Calculation = xlCalculationManual

Calculate
If Range("debtapplicable") = True Then

    Range("Sizing_Active") = 1
    Calculate
    projectnumber = Range("ProjectNumber")
    
    Do
      
       Range("Use_paste").Offset(projectnumber - 1, 0).Value = Range("Use_live").Value
          'Range("tax_paste").Offset(projectnumber - 1, 0).Value = Range("tax_live").Value
       Range("d_service_paste").Offset(projectnumber - 1, 0).Value = Range("d_service_live").Value
   
      Calculate
    
    Loop While Range("debt_delta") > 0.2 Or Range("Use_delta") > 0.2

 
         
    Range("SeniorDebtOptimalValue").Value = Range("SeniorDebtOptimalLive").Value
    
   
        Range("Sizing_Active") = 0
   
    
    Calculate
End If



'Range("tax_paste").Value = Range("tax_live").Value

Application.Calculation = calcmode

End Sub
