Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim l_Ret           As sNp_Shape
    Dim l_Direct        As Integer
    l_Direct = 0
    
    If g_Change_Move_Flag = 1 Then
        If _
            Target.Value = "U" Then
            Target.Interior.ColorIndex = g_GColor
            l_Direct = 4
        ElseIf _
            Target.Value = "D" And g_GameLevel <> 8 Then
            Target.Interior.ColorIndex = g_GColor
            g_GameLevelTmp = g_GameLevel
            g_GameLevel = 8
        ElseIf _
            Target.Value = "L" Then
            Target.Interior.ColorIndex = g_GColor
            l_Direct = 2
        ElseIf _
            Target.Value = "R" Then
            Target.Interior.ColorIndex = g_GColor
            l_Direct = 3
        ElseIf _
            Target.Value = "" Then
        End If
        
        If sNp_DirectionChange(l_Direct, g_CurrentShape, l_Ret) Then
        Else
        End If

        Target.Interior.ColorIndex = xlNone
        g_Bench.Activate
    Else
    End If
    
End Sub
