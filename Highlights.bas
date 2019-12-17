Attribute VB_Name = "Highlights"
Sub Highlighting()
Attribute Highlighting.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' Highlighting Macro
'
' Keyboard Shortcut: Ctrl+Shift+Q
'


    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
