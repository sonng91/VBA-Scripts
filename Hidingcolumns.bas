Attribute VB_Name = "Hidingcolumns"
Sub HideColumns()
Attribute HideColumns.VB_ProcData.VB_Invoke_Func = "X\n14"
'
' HideColumns Macro
'

    If Range("O1").Value = "Staged Count" Then
        columns("I:I").Delete
        columns("K:K").Delete
    End If

    columns("C:J").Hidden = True
    'Selection.EntireColumn.Hidden = True
    
    If Range("O1").Value = "Tracking Number" Then
        columns("C:L").Hidden = True
    End If

    If Range("N1").Value = "Tracking Number" Then
        columns("C:K").Hidden = True
    End If
    
End Sub

