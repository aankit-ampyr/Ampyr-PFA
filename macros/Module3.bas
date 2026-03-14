Attribute VB_Name = "Module3"
Sub Style_killer()
Dim N As Long, i As Long
With ActiveWorkbook
    
    N = .Styles.Count
    
    For i = N To 1 Step -1
        
        If Not .Styles(i).BuiltIn Then
            .Styles(i).Locked = False
            .Styles(i).Delete
        End If
        
    Next i
    
End With

End Sub
