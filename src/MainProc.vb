Public Sub sNp_Main()

    Dim l_TmpShapeBefore    As sNp_Shape
    Dim l_TmpShapeAfter     As sNp_Shape
    
    Dim l_Ret               As Boolean
    Dim l_TmpForAnyBoo      As Boolean
    
    Dim l_Counter           As Integer
    Dim l_EraserLine        As Integer
    Dim l_InnerCounter      As Integer
    Dim l_TmpForAnyInt1     As Integer
    Dim l_TmpForAnyInt2     As Integer
    
    Dim l_TmpForAnyStr      As String
    
    Dim l_TmpGameMarkRate   As Long
    
    'initialization
    g_GameLevel = 1
    g_GameLevelTmp = 1
    g_LevelMarks = 0
    g_TotalMarks = 0
    
    g_XBase = "E"
    g_YBase = "8"
    g_YView = "10"
    g_XWidth = 14
    g_YLength = 26
    g_PushBase.l_X = 7
    g_PushBase.l_Y = 1
    g_NextWindowBase.l_X = 19
    g_NextWindowBase.l_Y = 4
                
    g_ShapeDeadFlag = 2
    g_Change_Move_Flag = 1
    g_MergeFlag = False
    g_StateProc = True
    g_Pause = False
    
    Range("level") = CStr(g_GameLevel)
    Range("levelMark").Value = CStr(0)
    Range("totalMark").Value = CStr(0)
    Set g_NextWindow = Range("V11:Z13")
    Set g_WorkWindow = Range("E10:R33")
    Set g_Bench = Range("X21")
    g_Bench.Name = "Bench"
    g_NextWindow.Interior.ColorIndex = xlNone
    g_WorkWindow.Interior.ColorIndex = xlNone
    g_GColor = Range("color").Interior.ColorIndex
    l_TmpForAnyInt1 = 100
    l_TmpForAnyInt2 = 0
    
    sNp_KeySimulateConfig
    
NextLevel:
    'check in MainLoop
    Do
        If g_ShapeDeadFlag = 2 Then
            If _
                g_NextShape.l_Exist = 1 Then
                g_CurrentShape = g_NextShape
                g_NextShape = sNp_GetShapeByBase_ShapeAndMode(g_PushBase, Int((5 * Rnd) + 1), 1)
            Else
                g_CurrentShape = sNp_GetShapeByBase_ShapeAndMode(g_PushBase, Int((5 * Rnd) + 1), 1)
                g_NextShape = sNp_GetShapeByBase_ShapeAndMode(g_PushBase, Int((5 * Rnd) + 1), 1)
            End If
            l_Ret = sNp_RefreshNextWindow(g_NextShape)
        End If
        
        g_Change_Move_Flag = 1
        'D hangup process for level time.
        sNp_HangUp 33 - 4 * g_GameLevel
        g_Change_Move_Flag = 2
        
        'E@ check if the shape can fall down.
        If sNp_DirectionChange(1, g_CurrentShape, l_TmpShapeAfter) Then
            g_ShapeDeadFlag = 1
        Else
            If (g_MergeFlag) Then
                g_MergeFlag = False
                g_GameLevel = g_GameLevelTmp
                '-------------------------------------------------------------------------------------------------
                'getting g_leastY
                For l_Counter = 0 To 3
                    If g_CurrentShape.l_GFlag(l_Counter).l_Y < l_TmpForAnyInt1 Then
                        l_TmpForAnyInt1 = g_CurrentShape.l_GFlag(l_Counter).l_Y
                    End If
                    If g_CurrentShape.l_GFlag(l_Counter).l_Y > l_TmpForAnyInt2 Then
                        l_TmpForAnyInt2 = g_CurrentShape.l_GFlag(l_Counter).l_Y
                    End If
                Next l_Counter
                g_leastY = l_TmpForAnyInt1
                If g_leastY = 1 Then
                    GoTo GAMEOVER
                End If
                '-------------------------------------------------------------------------------------------------
                
                'checking if the player has make marks.
                l_TmpGameMarkRate = 0
                For l_Counter = l_TmpForAnyInt1 To l_TmpForAnyInt2
                    l_TmpForAnyBoo = True
                    For l_InnerCounter = 1 To g_XWidth
                        If sNp_GetRangeByMapG(l_InnerCounter, l_Counter).Interior.ColorIndex <> g_GColor Then
                            l_TmpForAnyBoo = False
                            Exit For
                        End If
                    Next l_InnerCounter
                    
                    If l_TmpForAnyBoo Then
                        Range(g_XBase & CStr(g_leastY + CLng(g_YBase) - 1) & ":" & Chr(Asc(g_XBase) + g_XWidth - 1) & CStr(l_Counter + g_YBase - 2)).Cut
                        g_leastY = g_leastY + 1
                        Application.ActiveSheet.Paste Range(g_XBase & CStr(g_leastY + g_YBase - 1))
                        g_WorkWindow.Borders(xlInsideVertical).LineStyle = xlDash
                        g_WorkWindow.Borders(xlInsideHorizontal).LineStyle = xlDash
                        Range("levelMark").Value = CStr(CLng(Range("levelMark").Value) + 10 + g_GameLevel * l_TmpGameMarkRate * 5)
                        Range("totalMark").Value = CStr(CLng(Range("totalMark").Value) + 10 + g_GameLevel * l_TmpGameMarkRate * 5)
                        l_TmpGameMarkRate = l_TmpGameMarkRate + 1
                        If CLng(Range("levelMark").Value) > 200 * g_GameLevel Then
                            If g_GameLevel = 8 Then
                                MsgBox ("CONGRATULATIONS!")
                                GoTo GAMEOVER
                            End If
                            Range("levelMark").Value = ""
                            g_GameLevel = CLng(Range("level").Value) + 1
                            Range("level") = CStr(g_GameLevel)
                            g_GameLevelTmp = g_GameLevel
                            'clear
                            g_NextWindow.Interior.ColorIndex = xlNone
                            g_WorkWindow.Interior.ColorIndex = xlNone

                            g_NextShape.l_Exist = 2
                            GoTo NextLevel
                        End If
                    End If
                Next l_Counter
                g_ShapeDeadFlag = 2
            End If
        End If
        While (g_Pause)
            'Pause
            DoEvents
        Wend
    Loop While (g_StateProc)
GAMEOVER:
    g_NextWindow.Interior.ColorIndex = xlNone
    g_WorkWindow.Interior.ColorIndex = xlNone
    Range("level") = ""
    Range("levelMark").Value = CStr(0)
    Range("totalMark").Value = CStr(0)
    MsgBox ("GAME OVER")
End Sub

