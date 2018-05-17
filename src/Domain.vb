'global viraiable's defination
'----------------------------------------
'baseline of arrowkey
Public g_Bench              As Range
Public g_NextWindow         As Range
Public g_WorkWindow         As Range

Public g_LevelMarks         As Long
Public g_TotalMarks         As Long

Public g_NextWindowBase     As sNp_Point

Public g_CurrentShape   As sNp_Shape

Public g_NextShape      As sNp_Shape

'1:alive   2:dead
Public g_ShapeDeadFlag      As Integer

'1:enable 2:disable
Public g_Change_Move_Flag   As Integer

Public g_standard       As Integer

Public g_leastY         As Integer

Public g_GameLevelTmp   As Integer

Public g_XBase          As String
Public g_YBase          As String
Public g_YView          As String
Public g_XWidth         As Integer
Public g_YLength        As Integer
Public g_PushBase       As sNp_Point

Public g_GameLevel      As Integer
Public g_IntervalTime   As Integer
Public g_GColor         As Integer
Public g_MergeFlag      As Boolean

Public g_StateProc      As Boolean
Public g_Pause          As Boolean

Public Type sNp_Point
    l_X                 As Integer
    l_Y                 As Integer
End Type

Public Type sNp_Shape
    '1:Exits  2:Not
    l_ShapeType         As Integer
    l_ShapeMode         As Integer
    l_Exist             As Integer
    l_GFlag(4)          As sNp_Point
End Type
'--------------------------------------------------
