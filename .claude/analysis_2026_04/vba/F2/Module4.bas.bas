Attribute VB_Name = "Module4"
Sub junior_sizing()

'If Range("Junior1_active") = True Then
    
    'Range("Junior1_paste").Value = Range("Junior1_start").Value
    
    'Do Until Range("Junior1_delta").Value < 1
        'Range("Junior1_paste").Value = Range("Junior1_copy").Value
    'Loop

'End If

If Range("Junior2_active") = True Then
    
    Range("Junior2_paste").Value = Range("Junior2_start").Value
    
    Do Until Range("Junior2_delta").Value < 1
        Range("Junior2_paste").Value = Range("Junior2_copy").Value
    Loop

End If

End Sub
