'Make effect that likes hangupping process.<currency>
Public Function sNp_HangUp(ByVal p_TimeInterval As Single)
    Dim l_TimeStart       As Single
    l_TimeStart = Timer
    Do
        DoEvents
    Loop While (Timer * 100 - l_TimeStart * 100 < p_TimeInterval)
End Function


'Simple method to response keyEvent in excel.<Excel,Impropriating Cursor>
'First u must code like this before this function transfered:
'Dim ws-> = This worksheet object.
'ws->.Range("given excel range without avail").activate

Public Sub sNp_KeySimulateConfig()

    Dim l_RangeUp               As Range
    Dim l_RangeDown             As Range
    Dim l_RangeLeft             As Range
    Dim l_RangeRight            As Range
    Dim l_TmpRangeArray(2)      As String
    Dim TmpFirst                As String
    Dim TmpSecond               As String
    
    TmpFirst = Strings.Split(g_Bench.AddressLocal, "$")(1)
    TmpSecond = Strings.Split(g_Bench.AddressLocal, "$")(2)
    
    Set l_RangeUp = g_Bench.Parent.Range(TmpFirst & CStr(CLng(TmpSecond) - 1))
    Set l_RangeDown = g_Bench.Parent.Range(TmpFirst & CStr(CLng(TmpSecond) + 1))
    
    Set l_RangeLeft = g_Bench.Parent.Range( _
        Strings.left(TmpFirst, CLng(Strings.Len(TmpFirst)) - 1) & Chr(Asc(Strings.right(Strings.Split(g_Bench.AddressLocal, "$")(1), 1)) - 1) & TmpSecond)
    Set l_RangeRight = g_Bench.Parent.Range( _
        Strings.left(TmpFirst, CLng(Strings.Len(TmpFirst)) - 1) & Chr(Asc(Strings.right(Strings.Split(g_Bench.AddressLocal, "$")(1), 1)) + 1) & TmpSecond)

    l_RangeUp.Name = "Up"
    l_RangeDown.Name = "Down"
    l_RangeLeft.Name = "Left"
    l_RangeRight.Name = "Right"
    
    g_Bench.Activate
    
End Sub

Public Function sNp_GetRangeByMapG( _
    ByVal p_X As Integer, _
    ByVal p_Y As Integer) As Range
    Set sNp_GetRangeByMapG = Range(Chr(Asc(g_XBase) + p_X - 1) & CStr(CLng(p_Y) + CLng(g_YBase) - 1))
End Function


Public Function sNp_RefreshNextWindow( _
    ByRef p_NextShape As sNp_Shape) As Boolean
    Dim l_TmpShape      As sNp_Shape
    Dim l_Counter       As Integer
    
    l_TmpShape = sNp_GetShapeByBase_ShapeAndMode(g_NextWindowBase, p_NextShape.l_ShapeType, p_NextShape.l_ShapeMode)
    g_NextWindow.Interior.ColorIndex = xlNone
    
    For l_Counter = 0 To 3
        sNp_GetRangeByMapG(l_TmpShape.l_GFlag(l_Counter).l_X, l_TmpShape.l_GFlag(l_Counter).l_Y).Interior.ColorIndex = g_GColor
    Next l_Counter
End Function

'p_downShape....bug. --> means p_Moving.
Public Function sNp_DirectionChange( _
    ByVal p_Direction As Integer, _
    ByRef p_DownShape As sNp_Shape, _
    ByRef p_RetShape As sNp_Shape) As Boolean
    
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim l_TmpShape      As sNp_Shape
    Dim l_Counter       As Integer
    Dim l_CounterInner  As Integer
    Dim l_TmpRange      As Range
    Dim l_BFlag         As Boolean
    
    '--------------------------------------------------------------------------------------------------------------------------------------------
    l_TmpShape = p_DownShape
    Select Case p_Direction
    Case 1 'Down
        l_TmpShape.l_GFlag(0).l_X = p_DownShape.l_GFlag(0).l_X
        l_TmpShape.l_GFlag(0).l_Y = p_DownShape.l_GFlag(0).l_Y + 1
    Case 2 'Left
        l_TmpShape.l_GFlag(0).l_X = p_DownShape.l_GFlag(0).l_X - 1
        l_TmpShape.l_GFlag(0).l_Y = p_DownShape.l_GFlag(0).l_Y
    Case 3 'Right
        l_TmpShape.l_GFlag(0).l_X = p_DownShape.l_GFlag(0).l_X + 1
        l_TmpShape.l_GFlag(0).l_Y = p_DownShape.l_GFlag(0).l_Y
    Case 4 'Change
        Select Case l_TmpShape.l_ShapeType
            Case 1
            Case 2
                l_TmpShape.l_ShapeMode = l_TmpShape.l_ShapeMode + 1
                If l_TmpShape.l_ShapeMode = 3 Then
                    l_TmpShape.l_ShapeMode = 1
                End If
            Case 3
                l_TmpShape.l_ShapeMode = l_TmpShape.l_ShapeMode + 1
                If l_TmpShape.l_ShapeMode = 5 Then
                    l_TmpShape.l_ShapeMode = 1
                End If
            Case 4
                l_TmpShape.l_ShapeMode = l_TmpShape.l_ShapeMode + 1
                If l_TmpShape.l_ShapeMode = 5 Then
                    l_TmpShape.l_ShapeMode = 1
                End If
            Case 5
                l_TmpShape.l_ShapeMode = l_TmpShape.l_ShapeMode + 1
                If l_TmpShape.l_ShapeMode = 5 Then
                    l_TmpShape.l_ShapeMode = 1
                End If
        End Select
    End Select
    
    l_TmpShape = sNp_GetShapeByBase_ShapeAndMode(l_TmpShape.l_GFlag(0), l_TmpShape.l_ShapeType, l_TmpShape.l_ShapeMode)
    
    '--------------------------------------------------------------------------------------------------------------------------------------------
    For l_Counter = 0 To 3
        l_BFlag = False
        Set l_TmpRange = sNp_GetRangeByMapG(l_TmpShape.l_GFlag(l_Counter).l_X, l_TmpShape.l_GFlag(l_Counter).l_Y)
        If _
            l_TmpRange.Interior.ColorIndex = g_GColor Then
            For l_CounterInner = 0 To 3
                If sNp_GetRangeByMapG(p_DownShape.l_GFlag(l_CounterInner).l_X, p_DownShape.l_GFlag(l_CounterInner).l_Y).AddressLocal = l_TmpRange.AddressLocal Then
                    l_BFlag = True
                Else
                End If
            Next l_CounterInner
            If l_BFlag Then
            Else
                g_MergeFlag = True
                sNp_DirectionChange = False
                Exit Function
            End If
        End If
        
        If _
            l_TmpShape.l_GFlag(l_Counter).l_Y > g_YLength Then
            sNp_DirectionChange = False
            g_MergeFlag = True
            Exit Function
        End If
        
        If _
            l_TmpShape.l_GFlag(l_Counter).l_X <= 0 Or l_TmpShape.l_GFlag(l_Counter).l_X > g_XWidth Or l_TmpShape.l_GFlag(l_Counter).l_Y <= 0 Then
            sNp_DirectionChange = False
            Exit Function
        End If
        
        
    Next l_Counter
    
    '--------------------------------------------------------------------------------------------------------------------------------------------
    For l_Counter = 0 To 3
        If p_DownShape.l_GFlag(l_Counter).l_Y >= 3 Then
            sNp_GetRangeByMapG(p_DownShape.l_GFlag(l_Counter).l_X, p_DownShape.l_GFlag(l_Counter).l_Y).Interior.ColorIndex = xlNone
        End If
    Next l_Counter
    For l_Counter = 0 To 3
        If l_TmpShape.l_GFlag(l_Counter).l_Y >= 3 Then
            sNp_GetRangeByMapG(l_TmpShape.l_GFlag(l_Counter).l_X, l_TmpShape.l_GFlag(l_Counter).l_Y).Interior.ColorIndex = g_GColor
        End If
    Next l_Counter
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    g_CurrentShape = l_TmpShape
    p_RetShape = l_TmpShape
    sNp_DirectionChange = True
    '--------------------------------------------------------------------------------------------------------------------------------------------
End Function

Public Function sNp_GetShapeByBase_ShapeAndMode( _
    ByRef p_Base As sNp_Point, _
    ByVal p_Type As Integer, _
    ByVal p_Mode As Integer) As sNp_Shape
    Dim l_Tmp   As sNp_Shape
    
    Select Case p_Type
    Case 1
        Select Case p_Mode
        Case 1
            l_Tmp.l_Exist = 1: l_Tmp.l_ShapeMode = p_Mode: l_Tmp.l_ShapeType = p_Type: l_Tmp.l_GFlag(0) = p_Base
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        End Select
    Case 2
        l_Tmp.l_Exist = 1: l_Tmp.l_ShapeMode = p_Mode: l_Tmp.l_ShapeType = p_Type: l_Tmp.l_GFlag(0) = p_Base
        Select Case p_Mode
        Case 1
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X + 2: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 2
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 2
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y - 1
        End Select
    Case 3
        l_Tmp.l_Exist = 1: l_Tmp.l_ShapeMode = p_Mode: l_Tmp.l_ShapeType = p_Type: l_Tmp.l_GFlag(0) = p_Base
        Select Case p_Mode
        Case 1
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 2
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y - 1
        Case 3
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 4
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y + 1
        End Select
    Case 4
        l_Tmp.l_Exist = 1: l_Tmp.l_ShapeMode = p_Mode: l_Tmp.l_ShapeType = p_Type: l_Tmp.l_GFlag(0) = p_Base
        Select Case p_Mode
        Case 1
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 2
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y - 1
        Case 3
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 4
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y + 1
        End Select
    Case 5
        l_Tmp.l_Exist = 1: l_Tmp.l_ShapeMode = p_Mode: l_Tmp.l_ShapeType = p_Type: l_Tmp.l_GFlag(0) = p_Base
        Select Case p_Mode
        Case 1
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 2
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y - 1
        Case 3
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X + 1: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y
        Case 4
            l_Tmp.l_GFlag(1).l_X = p_Base.l_X: l_Tmp.l_GFlag(1).l_Y = p_Base.l_Y - 1
            l_Tmp.l_GFlag(2).l_X = p_Base.l_X - 1: l_Tmp.l_GFlag(2).l_Y = p_Base.l_Y + 1
            l_Tmp.l_GFlag(3).l_X = p_Base.l_X: l_Tmp.l_GFlag(3).l_Y = p_Base.l_Y + 1
        End Select
    End Select
    sNp_GetShapeByBase_ShapeAndMode = l_Tmp
End Function

Public Sub sNp_TerminateProc()
    g_StateProc = False
End Sub

Public Sub sNp_PauseProc()
    If g_Pause = False Then
        g_Pause = True
    Else
        g_Pause = False
    End If
End Sub


