'$Dynamic
$ExeIcon:'./codeeditor.ico'
$Resize:On
DefLng A-Z

Const True = -1, False = 0

$If WIN Then
        Const FILE_SEPERATOR = "\"
$Else
                Const FILE_SEPERATOR = "/"
$End If

Type File
        As String Name, Path, Content, Cursors
        As _Unsigned Long TotalLines, CurrentCursor
        As Long ScrollOffset, FileID, SaveOffset, HorizontalScrollOffset
        As _Unsigned _Byte Opened, Saved
End Type
Dim Shared File(0) As File, CurrentFile As Long
Dim Shared Rope(0) As String, EmptyRopePoints As String
EmptyRopePoints = ""

Dim Shared As String Console

'-------- UI --------
Type UI
        As String Name, Label, Content, ParsedLabel, ParsedLabelUnderlined, ParsedContent, ParsedContentUnderlined
        As _Unsigned _Byte Visible, Type, Property, State, TotalStates
        As _Unsigned Integer Parent
        As Integer X, Y, W, H, __X, __Y, __W, __H
        As _Unsigned Long Response, Image, BG, FG, Key0_1, Key0_2, Key1_1, Key1_2, OnKey, Selected, MenuDialogWidth, MenuDialogHeight
        As Single Progress
        As Long Scroll, MaxScroll
End Type
Const UI_TYPE_Label = 1, UI_TYPE_Button = 2, UI_TYPE_Dialog = 3, UI_TYPE_ProgressBar = 4, UI_TYPE_MenuButton = 5, UI_TYPE_Image = 6, UI_TYPE_ScrollBar = 7, UI_TYPE_ToggleButton = 8, UI_TYPE_ListView = 9, UI_TYPE_TextView = 10, UI_TYPE_Frame = 11
Const UI_PROP_Center = 17
Const UI_KEY_Visible = 33, UI_KEY_Focus = 34, UI_KEY_Visible_And_Focus = 35
Dim Shared UI(0) As UI, UI_PARENT As _Unsigned Integer, UI_Focus_Parent As _Unsigned Integer, UI_Focus As _Unsigned Integer, UI_MouseWheel As Integer
Dim Shared As _Byte KeyCtrl, KeyAlt, KeyShift: Dim Shared KeyHit As Long
'--------------------

Dim Shared As String SaveFileQueue

Dim Shared Config$, Config_BackgroundColor As Long, Config_TextColor As Long, Config_CurrentWorkspace As String, Config_FileSeperator As String, Config_OpenFileSpeed As _Unsigned Long
NewConfig
CreateConfig
ReadConfigFile

'-------- Menu Bar --------
UI_MenuButton_File = UI_New("UI_MenuButton_File", UI_TYPE_MenuButton, 0, 0, 48, 16, " File ", ListStringFromString("\New File|Ctrl+N,\Open File|Ctrl+O,\Save File|Ctrl+S,Save \As File|Ctrl+Shift+S,\Close File|Ctrl+W,E\xit|Ctrl+Q"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 70, 100307, 102
UI_MenuButton_Workspace = UI_New("UI_MenuButton_Workspace", UI_TYPE_MenuButton, 48, 0, 88, 16, " Workspace ", ListStringFromString("\Open Workspace,\Save Workspace|Ctrl+Alt+S"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 87, 100307, 119
UI_MenuButton_View = UI_New("UI_MenuButton_View", UI_TYPE_MenuButton, 136, 0, 48, 16, " View ", ListStringFromString("Go to \File,Go to \Line|Ctrl+G,\Toggle Side Pane|Alt+T,Toggle Summary Bar|Alt+B"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 86, 100307, 118
UI_MenuButton_Cursors = UI_New("UI_MenuButton_Cursors", UI_TYPE_MenuButton, 184, 0, 72, 16, " Cursors ", ListStringFromString("Go to \Next Cursor|Alt+X,Set Cursor to Center|F8"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 67, 100307, 99
Dim Shared As Integer FileChangeDialog
'--------------------------
'-------- Scroll Bar --------
UI_ScrollBar_Text = UI_New("UI_ScrollBar_Text", UI_TYPE_ScrollBar, -8, 17, -1, -16, "", "")
Dim Shared As Integer Summary_Bar
'----------------------------
'-------- Status Bar --------
Dim Shared As Integer UI_Label_CursorPosition
UI_ToggleButton_LineSeperator = UI_New("UI_ToggleButton_LineSeperator", UI_TYPE_ToggleButton, 16, -16, 48, 16, "", ListStringFromString("[CRLF],[  LF],[CR  ]"))
Select Case Config_FileSeperator
        Case Chr$(13) + Chr$(10): UI(UI_ToggleButton_LineSeperator).State = 0
        Case Chr$(10): UI(UI_ToggleButton_LineSeperator).State = 1
        Case Chr$(13): UI(UI_ToggleButton_LineSeperator).State = 2
End Select
UI_Set_BG _RGB32(0, 191, 0)
UI_Label_CursorPosition = UI_New("UI_Label_CursorPosition", UI_TYPE_Label, 96, -16, 128, 16, "Cursor: 0 0", "")
'----------------------------
'-------- Draw Text --------
Dim Shared As Long VerticalLinesVisible, HorizontalCharsVisible
'---------------------------
'-------- Side Pane --------
Dim Shared As Integer UI_Side_Pane
UI_Side_Pane = UI_New("UI_Side_Pane", UI_TYPE_Frame, 0, 17, 16, -16, "", "")
UI_Set_BG _RGB32(47)
'---------------------------
'-------- Open File Dialog --------
Dim Shared As Integer UI_Dialog_OpenFile
OpenFileDialog 'First Run
'----------------------------------
'-------- Save As Dialog --------
Dim Shared As Integer UI_Dialog_SaveFileAs
SaveFileAsDialog 'First Run
'--------------------------------
'-------- Go to Line Dialog --------
Dim Shared As Integer UI_Dialog_GoToLine
GoToLineDialog 'First Run
'-----------------------------------

Dim Shared As Long MainScreen, tmpScreen
MainScreen = _NewImage(960, 540, 32)
Screen MainScreen
Do Until _ScreenExists: Loop
While _Resize: Wend

_AcceptFileDrop

Color Config_TextColor, 0

If _CommandCount Then
        For I = 1 To _CommandCount
                If _FileExists(Command$(I)) Then INFILE$ = Command$(I) Else INFILE$ = _StartDir$ + FILE_SEPERATOR + Command$(I)
                OpenFile INFILE$
        Next I
Else
        If _FileExists(Config_CurrentWorkspace) Then OpenWorkspace Config_CurrentWorkspace
End If

Dim Shared As Long TextOffsetX, TextOffsetY: TextOffsetY = 16

Do
        If _Resize Then
                tmpW = _ResizeWidth: tmpH = _ResizeHeight
                If tmpW > 0 And tmpH > 0 Then tmpScreen = MainScreen: MainScreen = _NewImage(tmpW, tmpH, 32): Screen MainScreen: _FreeImage tmpScreen
        End If
        If _TotalDroppedFiles Then
                For I = 1 To _TotalDroppedFiles
                        If _FileExists(_DroppedFile$(I)) Then INFILE$ = _DroppedFile$(I) Else INFILE$ = _StartDir$ + FILE_SEPERATOR + _DroppedFile$(I)
                        OpenFile INFILE$
                Next I
                _FinishDrop
        End If
        Cls , Config_BackgroundColor
        Color Config_TextColor, 0
        _Limit 60
        CurrentFile = Clamp(LBound(File), CurrentFile, UBound(File))
        If CurrentFile Then
                If File(CurrentFile).Saved = 0 Then
                        SetTitle "*" + File(CurrentFile).Name + " - Code Editor"
                Else
                        SetTitle File(CurrentFile).Name + " - Code Editor"
                End If
        Else
                SetTitle "Code Editor"
        End If
        UI_MouseWheel = 0
        While _MouseInput
                UI_MouseWheel = UI_MouseWheel + _MouseWheel
                If UI_Focus = 0 And _MouseX > UI(UI_Side_Pane).W Then File(CurrentFile).ScrollOffset = Clamp(1, File(CurrentFile).ScrollOffset + 3 * _MouseWheel, File(CurrentFile).TotalLines)
        Wend
        If UI_MouseWheel Then MoveScrollToCursor -1

        OpenFileTasks: SaveFileTasks

        K$ = InKey$: KeyHit = _KeyHit
        KeyShift = _KeyDown(100303) Or _KeyDown(100304)
        KeyCtrl = _KeyDown(100305) Or _KeyDown(100306)
        KeyAlt = _KeyDown(100307) Or _KeyDown(100308)

        DrawMenuBar
        DrawStatusBar
        If CurrentFile Then
                TextOffsetX = UI(UI_Side_Pane).W + 8 * ceil(Log(Max(1, File(CurrentFile).TotalLines)) / Log(10) + 2)
                MoveScrollToCursor 0
                DrawText
                If Summary_Bar Then DrawSummaryBar
                If UI(UI_Side_Pane).W >= 128 Then DrawSidePane
        End If
        If UI(UI_Dialog_OpenFile).Visible Then
                OpenFileDialog
        ElseIf UI(UI_Dialog_SaveFileAs).Visible Then
                SaveFileAsDialog
        ElseIf UI(UI_Dialog_GoToLine).Visible Then
                GoToLineDialog
        ElseIf UI_Focus = 0 And CurrentFile Then 'Handle Text Area
                If UI(UI_ScrollBar_Text).Response Then File(CurrentFile).ScrollOffset = UI(UI_ScrollBar_Text).Scroll
                If _MouseButton(1) And MouseInBox(0, 16, _Width - 9, _Height - 17) Then
                        If KeyAlt Then AddCursor _SHR(_MouseX - TextOffsetX, 3) + File(CurrentFile).HorizontalScrollOffset, _SHR(_MouseY - TextOffsetY, 4) + File(CurrentFile).ScrollOffset Else SetCursor _SHR(_MouseX - TextOffsetX, 3) + File(CurrentFile).HorizontalScrollOffset, _SHR(_MouseY - TextOffsetY, 4) + File(CurrentFile).ScrollOffset
                End If
                If KeyAlt And Len(File(CurrentFile).Cursors) = 8 Then
                        If KeyShift And Left$(Console, 1) <> "-" Then Console = "-" + Console
                        If InRange(48, KeyHit, 57) Then Console = Console + Chr$(KeyHit)
                Else
                        If LastKeyAlt And Len(File(CurrentFile).Cursors) = 8 Then
                                SetCursor GetCursorX, GetCursorY + Val(Console)
                        Else
                                Console = ""
                        End If
                End If
                LastKeyAlt = KeyAlt
                UI(UI_ScrollBar_Text).MaxScroll = File(CurrentFile).TotalLines
                UI(UI_ScrollBar_Text).Scroll = File(CurrentFile).ScrollOffset
                For I = 1 To Len(File(CurrentFile).Cursors) Step 8
                        CursorX = CVL(Mid$(File(CurrentFile).Cursors, I, 4))
                        CursorY = CVL(Mid$(File(CurrentFile).Cursors, I + 4, 4))
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        If RopeI = 0 Then _Continue
                        BoundCursor = False
                        If Len(K$) Then
                                Select Case Asc(K$)
                                        Case 8: DeleteText 1, CursorX, CursorY: BoundCursor = True
                                        Case 9: InsertText Space$(8), CursorX, CursorY: BoundCursor = True
                                        Case 13: S$ = Left$(Rope(RopeI), Len(Rope(RopeI)) - Len(LTrim$(Rope(RopeI))))
                                                If CursorX = Len(Rope(RopeI)) + 1 Then
                                                        InsertLine CursorY + 1
                                                        CursorX = Len(S$) + 1
                                                        CursorY = CursorY + 1
                                                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                                                        Rope(RopeI) = S$
                                                Else
                                                        InsertLine CursorY + 1
                                                        T$ = Mid$(Rope(RopeI), CursorX)
                                                        Rope(RopeI) = Left$(Rope(RopeI), CursorX - 1)
                                                        CursorY = CursorY + 1: CursorX = Len(S$) + 1
                                                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                                                        Rope(RopeI) = S$ + T$
                                                End If
                                                File(CurrentFile).HorizontalScrollOffset = 1
                                        Case 32 To 126: If KeyCtrl = 0 And KeyAlt = 0 Then InsertText K$, CursorX, CursorY: BoundCursor = True
                                End Select
                        End If
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        ShowCursor = Timer(0.1) - Int(Timer(0.1)) < 0.5
                        Select EveryCase KeyHit
                                Case 19200 'Left
                                        If KeyCtrl Then
                                                File(CurrentFile).HorizontalScrollOffset = Max(1, File(CurrentFile).HorizontalScrollOffset - 3)
                                        Else
                                                CursorX = Max(CursorX - 1, 1)
                                        End If
                                Case 19712 'Right
                                        If KeyCtrl Then
                                                File(CurrentFile).HorizontalScrollOffset = Min(File(CurrentFile).HorizontalScrollOffset + 3, Len(Rope(RopeI)) + 1)
                                        Else
                                                CursorX = Min(CursorX + 1, Len(Rope(RopeI)) + 1)
                                        End If
                                Case 18432 'Up
                                        If KeyCtrl Then
                                                File(CurrentFile).ScrollOffset = Max(1, File(CurrentFile).ScrollOffset - 1)
                                        Else
                                                CursorY = Max(CursorY - 1, 1)
                                                If CursorY < File(CurrentFile).ScrollOffset Then File(CurrentFile).ScrollOffset = CursorY
                                        End If
                                Case 20480 'Down
                                        If KeyCtrl Then
                                                File(CurrentFile).ScrollOffset = Min(File(CurrentFile).ScrollOffset + 1, File(CurrentFile).TotalLines)
                                        Else
                                                CursorY = Min(CursorY + 1, File(CurrentFile).TotalLines)
                                                If CursorY > File(CurrentFile).ScrollOffset + VerticalLinesVisible Then File(CurrentFile).ScrollOffset = Max(1, CursorY - VerticalLinesVisible)
                                        End If
                                Case 18688 'PgUp
                                        CursorY = Max(1, CursorY - VerticalLinesVisible)
                                        If CursorY < File(CurrentFile).ScrollOffset Then File(CurrentFile).ScrollOffset = CursorY
                                Case 20736 'PgDn
                                        CursorY = Min(CursorY + VerticalLinesVisible, File(CurrentFile).TotalLines)
                                        If CursorY > File(CurrentFile).ScrollOffset + VerticalLinesVisible Then File(CurrentFile).ScrollOffset = Max(1, CursorY - VerticalLinesVisible)
                                Case 19200, 19712, 18432, 20480, 18688, 19712: MoveScrollToCursor -1
                                Case 18176 'Home
                                        If KeyCtrl Then
                                                CursorX = 1: CursorY = 1
                                                MoveScrollToCursor 1
                                        Else
                                                CursorX = 1
                                        End If
                                        File(CurrentFile).HorizontalScrollOffset = 1
                                Case 20224 'End
                                        If KeyCtrl Then
                                                CursorX = 1
                                                CursorY = File(CurrentFile).TotalLines
                                                MoveScrollToCursor Max(1, File(CurrentFile).TotalLines - VerticalLinesVisible)
                                        Else
                                                CursorX = Len(Rope(RopeI)) + 1
                                        End If
                                        File(CurrentFile).HorizontalScrollOffset = Max(1, Len(Rope(RopeI)) + 1 - _SHR(HorizontalCharsVisible, 1))
                                Case 21248 'Delete
                                        If KeyCtrl Then
                                                DeleteLine CursorY
                                                File(CurrentFile).HorizontalScrollOffset = 1
                                                BoundCursor = True
                                        Else
                                                If CursorX = Len(Rope(RopeI)) + 1 And CursorY < File(CurrentFile).TotalLines Then
                                                        Rope(RopeI) = Rope(RopeI) + Rope(CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY + 1, 2) - 3, 4)))
                                                        DeleteLine CursorY + 1
                                                Else
                                                        DeleteText 1, CursorX + 1, CursorY
                                                End If
                                        End If
                                Case 86, 118: If KeyCtrl Then
                                                InsertText _Clipboard$, CursorX, CursorY
                                        End If
                        End Select
                        If BoundCursor Then
                                CursorX = Clamp(1, CursorX, Len(Rope(RopeI)) + 1)
                                CursorY = Clamp(1, CursorY, File(CurrentFile).TotalLines)
                                ShowCursorTime = Timer(0.1) + 1
                        End If
                        Mid$(File(CurrentFile).Cursors, I, 4) = MKL$(CursorX)
                        Mid$(File(CurrentFile).Cursors, I + 4, 4) = MKL$(CursorY)
                        If I = 1 Then UI(UI_Label_CursorPosition).Label = "Cursor:" + Str$(CursorX) + Str$(CursorY)
                        'Display Cursor
                        If ShowCursor = 0 And Timer(0.1) > ShowCursorTime Then _Continue
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        X = TextOffsetX + _SHL(Min(CursorX, Len(Rope(RopeI)) + 1) - File(CurrentFile).HorizontalScrollOffset, 3)
                        Y = TextOffsetY + _SHL(CursorY - File(CurrentFile).ScrollOffset, 4)
                        Line (X, Y)-(X + 7, Y + 15), -1, B
                Next I
                Select Case KeyHit
                        Case 9: If KeyCtrl Then CurrentFile = ClampCycle(LBound(File), CurrentFile + 1, UBound(File)) 'Ctrl Tab
                        Case 16384: 'F6
                                CurrentFile = ClampCycle(LBound(File), CurrentFile + 1, UBound(File))
                End Select
        End If
        UI_Draw: _Display

        'New File
        If UI(UI_MenuButton_File).Response = 1 Or (KeyCtrl And (KeyHit = 78 Or KeyHit = 110)) Then NewFile
        'Open File
        If UI(UI_MenuButton_File).Response = 2 Or (KeyCtrl And (KeyHit = 79 Or KeyHit = 111)) Then UI(UI_Dialog_OpenFile).Visible = -1: UI_Focus = UI_Dialog_OpenFile
        'Save File
        If UI(UI_MenuButton_File).Response = 3 Or (KeyCtrl And KeyShift = 0 And KeyAlt = 0 And (KeyHit = 83 Or KeyHit = 115)) Then SaveFile CurrentFile
        'Save As File
        If UI(UI_MenuButton_File).Response = 4 Or (KeyCtrl And KeyShift And (KeyHit = 83 Or KeyHit = 115)) Then UI(UI_Dialog_SaveFileAs).Visible = -1: UI_Focus = UI_Dialog_SaveFileAs
        'Close File
        If UI(UI_MenuButton_File).Response = 5 Or (KeyCtrl And (KeyHit = 87 Or KeyHit = 119)) Then CloseFile CurrentFile
        'Exit
        If UI(UI_MenuButton_File).Response = 6 Or (KeyCtrl And (KeyHit = 81 Or KeyHit = 113)) Then Exit Do
        'Open Workspace
        If UI(UI_MenuButton_Workspace).Response = 1 Then OpenWorkspace "codeeditor_workspace"
        'Save Workspace
        If UI(UI_MenuButton_Workspace).Response = 2 Or (KeyCtrl And KeyAlt And (KeyHit = 83 Or KeyHit = 115)) Then SaveWorkspace "codeeditor_workspace"
        'Go to File
        If UI(UI_MenuButton_View).Response = 1 Then UI(FileChangeDialog).Visible = -1: UI_Focus = FileChangeDialog
        'Go to Line
        If UI(UI_MenuButton_View).Response = 2 Or (KeyCtrl And (KeyHit = 71 Or KeyHit = 103)) Then UI(UI_Dialog_GoToLine).Visible = -1: UI_Focus = UI_Dialog_GoToLine
        'Toggle Side Pane
        If UI(UI_MenuButton_View).Response = 3 Or (KeyAlt And (KeyHit = 84 Or KeyHit = 116)) Then UI(UI_Side_Pane).W = 272 - UI(UI_Side_Pane).W
        'Toggle Summary Bar
        If UI(UI_MenuButton_View).Response = 4 Or (KeyAlt And (KeyHit = 66 Or KeyHit = 98)) Then Summary_Bar = 1 - Summary_Bar
        'Go to Next Cursor
        If UI(UI_MenuButton_Cursors).Response = 1 Or UI(UI_Label_CursorPosition).Response Or (KeyAlt And (KeyHit = 88 Or KeyHit = 120)) Then
                File(CurrentFile).CurrentCursor = ClampCycle(1, File(CurrentFile).CurrentCursor + 1, _SHR(Len(File(CurrentFile).Cursors), 3))
                File(CurrentFile).ScrollOffset = CVL(Mid$(File(CurrentFile).Cursors, _SHL(File(CurrentFile).CurrentCursor, 3) - 3, 4))
                UI_Focus = 0
        End If
        'Set Cursor to Center
        If UI(UI_MenuButton_Cursors).Response = 2 Or (KeyHit = 16896) Then
                SetCursor 1, File(CurrentFile).ScrollOffset + _SHR(VerticalLinesVisible, 1)
        End If
        'Line Seperator Button
        Select Case UI(UI_ToggleButton_LineSeperator).Response
                Case 1: Config_FileSeperator = Chr$(13) + Chr$(10)
                Case 2: Config_FileSeperator = Chr$(10)
                Case 3: Config_FileSeperator = Chr$(13)
        End Select

        If _Exit Then Exit Do
Loop
WriteConfigFile
System

'------- Cursor Management --------
Sub AddCursor (X As Long, Y As Long)
        For I = 1 To Len(File(CurrentFile).Cursors) Step 8
                CX = CVL(Mid$(File(CurrentFile).Cursors, I, 4))
                CY = CVL(Mid$(File(CurrentFile).Cursors, I + 4, 4))
                If CX = X And CY = Y Then Exit Sub
        Next I
        File(CurrentFile).Cursors = File(CurrentFile).Cursors + MKL$(Max(1, X)) + MKL$(Clamp(1, Y, File(CurrentFile).TotalLines))
End Sub
Sub SetCursor (X As Long, Y As Long)
        File(CurrentFile).Cursors = MKL$(Max(1, X)) + MKL$(Clamp(1, Y, File(CurrentFile).TotalLines))
End Sub
Function GetCursorX&
        GetCursorX = CVL(Left$(File(CurrentFile).Cursors, 4))
End Function
Function GetCursorY&
        GetCursorY = CVL(Mid$(File(CurrentFile).Cursors, 5, 4))
End Function
Sub MoveScrollToCursor (Y As Long)
        Static DestY As Long
        If Y Then DestY = Y
        If Y = -1 Then DestY = 0
        If DestY = 0 Then Exit Sub
        D = DestY - File(CurrentFile).ScrollOffset
        File(CurrentFile).ScrollOffset = File(CurrentFile).ScrollOffset + Sgn(D) * CeilDiv2(Abs(D), 2)
        If D = 0 Then DestY = 0
End Sub
'----------------------------------

'-------- Rope Management --------
Sub InsertText (T$, CursorX As Long, CursorY As Long)
        If InStr(T$, Chr$(10)) Or InStr(T$, Chr$(13)) Then
        Else
                RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                Rope(RopeI) = Left$(Rope(RopeI), CursorX - 1) + T$ + Mid$(Rope(RopeI), CursorX)
                CursorX = CursorX + Len(T$)
                UpdateRope RopeI
        End If
        File(CurrentFile).Saved = 0
End Sub
Sub InsertLine (CursorY As Long)
        Dim As Long RopeI
        File(CurrentFile).TotalLines = File(CurrentFile).TotalLines + 1
        RopeI = GetNewRopePointer
        File(CurrentFile).Content = Left$(File(CurrentFile).Content, _SHL(CursorY - 1, 2)) + MKL$(RopeI) + Mid$(File(CurrentFile).Content, _SHL(CursorY - 1, 2) + 1)
        UpdateRope RopeI
        File(CurrentFile).Saved = 0
End Sub
Sub DeleteText (Count As Long, CursorX As Long, CursorY As Long)
        For I = 1 To Count
                If CursorX = 1 Then
                        If CursorY = 1 Then Exit Sub
                        T$ = Rope(CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4)))
                        DeleteLine CursorY: CursorY = CursorY - 1
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        CursorX = Len(Rope(RopeI)) + 1
                        Rope(RopeI) = Rope(RopeI) + T$
                Else
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        Rope(RopeI) = Left$(Rope(RopeI), CursorX - 2) + Mid$(Rope(RopeI), CursorX)
                        UpdateRope RopeI
                        CursorX = CursorX - 1
                End If
        Next I
        File(CurrentFile).Saved = 0
End Sub
Sub DeleteLine (CursorY As Long)
        EmptyRopePoints = EmptyRopePoints + Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4)
        File(CurrentFile).Content = Left$(File(CurrentFile).Content, _SHL(CursorY - 1, 2)) + Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) + 1)
        File(CurrentFile).TotalLines = File(CurrentFile).TotalLines - 1
        File(CurrentFile).Saved = 0
End Sub
Sub UpdateRope (I As Long) Static
        Dim As String ParsedRope
        ParsedRope = ""
        For J = 1 To Len(Rope(I))
                B~%% = Asc(Rope(I), J)
                Select Case B~%%
                        Case 9: ParsedRope = ParsedRope + Space$(8)
                        Case Else: ParsedRope = ParsedRope + Chr$(B~%%)
        End Select: Next J
        Rope(I) = ParsedRope
        ParsedRope = ""
End Sub
'---------------------------------

Sub OpenFileDialog
        Static As Integer UI_Dialog_OpenFile_ListView_Dir, UI_Dialog_OpenFile_TextView_Dir, UI_Dialog_OpenFile_Button_Open, UI_Dialog_OpenFile_Button_Cancel
        Static CurrentDirectory As String
        If UI_Dialog_OpenFile = 0 Then
                UI_Dialog_OpenFile = UI_New("UI_Dialog_OpenFile", UI_TYPE_Dialog, 0, 0, -128, -128, "Open File", "")
                UI_Set_FG -1
                UI(UI_Dialog_OpenFile).Visible = 0
                UI_PARENT = UI_Dialog_OpenFile
                UI_Dialog_OpenFile_ListView_Dir = UI_New("UI_Dialog_OpenFile_ListView_Dir", UI_TYPE_ListView, 16, 64, -16, -32, "", "")
                UI_Set_BG _RGB32(0, 191, 0)
                UI_Dialog_OpenFile_TextView_Dir = UI_New("UI_Dialog_OpenFile_TextView_Dir", UI_TYPE_TextView, 16, 40, -16, 16, "", "")
                UI_Dialog_OpenFile_Button_Open = UI_New("UI_Dialog_OpenFile_Button_Open", UI_TYPE_Button, -128, -24, 48, 16, " Open ", "")
                UI_Dialog_OpenFile_Button_Cancel = UI_New("UI_Dialog_OpenFile_Button_Cancel", UI_TYPE_Button, -64, -24, 48, 16, "Cancel", "")
                UI_PARENT = 0
                CurrentDirectory = _StartDir$
                Exit Sub
        End If
        Update = 0
        If UI_Focus = 0 Then UI(UI_Dialog_OpenFile).Visible = 0
        Select Case UI(UI_Dialog_OpenFile_TextView_Dir).Response
                Case 13: If _DirExists(UI(UI_Dialog_OpenFile_TextView_Dir).Content) Then
                                CurrentDirectory = UI(UI_Dialog_OpenFile_TextView_Dir).Content
                                Update = 1
                        ElseIf _FileExists(UI(UI_Dialog_OpenFile_TextView_Dir).Content) Then
                                OpenFile UI(UI_Dialog_OpenFile_TextView_Dir).Content
                                CurrentDirectory = UI(UI_Dialog_OpenFile_TextView_Dir).Content
                                Update = 1
                                UI(UI_Dialog_OpenFile).Visible = 0
                                UI_Focus = 0
                        End If
        End Select
        If Len(UI(UI_Dialog_OpenFile_ListView_Dir).ParsedContent) = 0 Then Update = 1
        Select Case UI(UI_Dialog_OpenFile_ListView_Dir).Selected
                Case 0
                Case 1
                        CurrentDirectory = PathBack$(CurrentDirectory)
                        Update = 1
                Case Else
                        If _DirExists(CurrentDirectory + ListStringGet(UI(UI_Dialog_OpenFile_ListView_Dir).ParsedContent, UI(UI_Dialog_OpenFile_ListView_Dir).Selected)) Then
                                CurrentDirectory = CurrentDirectory + ListStringGet(UI(UI_Dialog_OpenFile_ListView_Dir).ParsedContent, UI(UI_Dialog_OpenFile_ListView_Dir).Selected)
                                Update = 1
                        End If
        End Select
        If UI(UI_Dialog_OpenFile_Button_Open).Response Then
                OpenFile CurrentDirectory + ListStringGet(UI(UI_Dialog_OpenFile_ListView_Dir).ParsedContent, UI(UI_Dialog_OpenFile_ListView_Dir).Selected)
                Update = 1
                UI(UI_Dialog_OpenFile).Visible = 0
                UI_Focus = 0
        End If
        If UI(UI_Dialog_OpenFile_Button_Cancel).Response Then
                UI(UI_Dialog_OpenFile).Visible = 0
                UI_Focus = 0
        End If
        If Update Then
                UI(UI_Dialog_OpenFile_ListView_Dir).ParsedContent = ListStringAppend(ListStringFromString(".."), GetDirList$(CurrentDirectory))
                UI(UI_Dialog_OpenFile_ListView_Dir).Selected = 0
                If UI_Focus <> UI_Dialog_OpenFile_TextView_Dir Then UI(UI_Dialog_OpenFile_TextView_Dir).Content = CurrentDirectory
        End If
End Sub
Sub SaveFileAsDialog
        If UI_Dialog_SaveFileAs = 0 Then
                UI_Dialog_SaveFileAs = UI_New("UI_Dialog_SaveFileAs", UI_TYPE_Dialog, 0, 0, -128, -128, "Save As File", "")
                UI_Set_FG -1
                UI(UI_Dialog_SaveFileAs).Visible = 0
                UI_PARENT = UI_Dialog_SaveFileAs
                UI_PARENT = 0
                Exit Sub
        End If
End Sub
Sub GoToLineDialog
        Static As Integer TextView, Button_Ok, Button_Cancel
        If UI_Dialog_GoToLine = 0 Then
                UI_Dialog_GoToLine = UI_New("UI_Dialog_GoToLine", UI_TYPE_Dialog, 0, 0, 128, 64, "Go to Line", "")
                UI_Set_FG -1
                UI(UI_Dialog_GoToLine).Visible = 0
                UI_PARENT = UI_Dialog_GoToLine
                TextView = UI_New("UI_Dialog_GoToLine_TextView", UI_TYPE_TextView, 8, 16, -8, 16, "", "")
                Button_Ok = UI_New("UI_Dialog_GoToLine_Button_Ok", UI_TYPE_Button, 8, 32, 48, 16, "  Ok  ", "")
                Button_Cancel = UI_New("UI_Dialog_GoToLine_Button_Cancel", UI_TYPE_Button, -64, 32, 48, 16, "Cancel", "")
                UI_PARENT = 0
                Exit Sub
        End If
        If UI_Focus = 0 Then UI(UI_Dialog_GoToLine).Visible = 0: Exit Sub
        UI_Focus = TextView
        If UI(Button_Ok).Response Then SetCursor 1, Val(UI(TextView).Content): File(CurrentFile).ScrollOffset = Val(UI(TextView).Content): UI(UI_Dialog_GoToLine).Visible = 0: UI_Focus = 0
        If UI(Button_Cancel).Response Then UI(UI_Dialog_GoToLine).Visible = 0: UI_Focus = 0
End Sub
Sub DrawMenuBar
        Static As Integer ListView, ScrollBar
        If FileChangeDialog = 0 Then
                FileChangeDialog = UI_New("UI_Dialog_FilesList", UI_TYPE_Dialog, -256, 16, 256, 32, "Files", "")
                UI_Set_FG -1
                UI(FileChangeDialog).Property = 0
                UI(FileChangeDialog).Visible = 0
                UI_PARENT = FileChangeDialog
                ListView = UI_New("UI_Dialog_FilesList_ListView", UI_TYPE_ListView, 8, 16, -16, -16, "", "")
                ScrollBar = UI_New("UI_Dialog_FilesList_ScrollBar", UI_TYPE_ScrollBar, -16, 16, -8, -16, "", "")
                UI_PARENT = 0
        End If
        If UI(FileChangeDialog).Visible Then
                If UI_Focus = 0 Then UI(FileChangeDialog).Visible = 0
                If UI(ScrollBar).Response Then UI(ListView).Scroll = UI(ScrollBar).Scroll
                UI(FileChangeDialog).H = _SHL(2 + Clamp(1, UBound(File), 4), 4)
                UI(ListView).ParsedContent = ListStringNew
                For I = 1 To UBound(File): ListStringAdd UI(ListView).ParsedContent, File(I).Name: Next I
                UI(ScrollBar).MaxScroll = UBound(File)
                UI(ScrollBar).Scroll = UI(ListView).Scroll
                If UI(ListView).Response Then CurrentFile = Clamp(1, UI(ListView).Response, UBound(File)): UI(FileChangeDialog).Visible = 0
        End If
        Line (0, 0)-(_Width - 1, 16), _RGB32(0, 63, 127), BF
        If CurrentFile Then
                _PrintString (_Width - 8 * Len(File(CurrentFile).Name) - 16, 0), "[" + File(CurrentFile).Name + "]"
                If MouseInBox(_Width - 8 * Len(File(CurrentFile).Name) - 16, 0, _Width, 16) And _MouseButton(1) Then UI(FileChangeDialog).Visible = -1: UI_Focus = FileChangeDialog: WaitForMouseButtonRelease
        End If
End Sub
Sub DrawSummaryBar

End Sub
Sub DrawStatusBar
        Line (0, _Height - 16)-(_Width - 1, _Height - 1), _RGB32(0, 63, 127), BF
        If KeyCtrl Then T$ = " Ctrl "
        If KeyAlt Then T$ = T$ + " Alt "
        If KeyShift Then T$ = T$ + " Shift "
        T$ = T$ + Console
        _PrintString (_Width - _SHL(Len(T$), 3), _Height - 16), T$
End Sub
Sub DrawSidePane
        Static FilesList$, ScrollOffset, LastCurrentFile
        If FilesList$ = "" Or LastCurrentFile <> CurrentFile Then FilesList$ = GetDirList$(PathBack$(File(CurrentFile).Path))
        If _MouseX < UI(UI_Side_Pane).W Then ScrollOffset = ScrollOffset + UI_MouseWheel
        J = 1: For I = ScrollOffset To Min(ScrollOffset + VerticalLinesVisible, ListStringLength(FilesList$)): _PrintString (0, 16 + _SHL(J, 4)), ListStringGet(FilesList$, I): J = J + 1: Next I
        LastCurrentFile = CurrentFile
End Sub
Sub DrawText
        VerticalLinesVisible = _SHR(_Height - 32, 4) - 1
        HorizontalCharsVisible = _SHR(_Width - TextOffsetX, 3) - Summary_Bar * 8
        J = TextOffsetY
        For I = File(CurrentFile).ScrollOffset To Min(File(CurrentFile).ScrollOffset + VerticalLinesVisible, File(CurrentFile).TotalLines)
                If InRange(1, I, File(CurrentFile).TotalLines) = 0 Then _Continue
                K = CVL(Mid$(File(CurrentFile).Content, I * 4 - 3, 4))
                If KeyAlt And GetCursorY <> I And Len(File(CurrentFile).Cursors) = 8 Then L~& = Abs(I - GetCursorY) Else L~& = I
                L$ = _Trim$(Str$(L~&))
                T$ = Mid$(Rope(K), File(CurrentFile).HorizontalScrollOffset)
                L$ = Space$(_SHR(TextOffsetX, 3) - Len(L$) - 1) + L$ + Chr$(179) + Left$(T$, HorizontalCharsVisible)
                _PrintString (0, J), L$
                J = J + 16
        Next I
End Sub

'-------- Config Files --------
Sub ReadConfigFile
        F = FreeFile
        If _FileExists("config.ini") = 0 Then Exit Sub
        Open "config.ini" For Binary As #F
        Config$ = String$(LOF(F), 0)
        Get #F, , Config$
        Close #F
        ParseConfig
End Sub
Sub NewConfig
        Config$ = MapNew
        Config_BackgroundColor = _RGB32(32)
        Config_TextColor = _RGB32(255)
        Config_CurrentWorkspace = "codeeditor_workspace"
        Config_FileSeperator = Chr$(13) + Chr$(10)
        Config_OpenFileSpeed = 256
End Sub
Sub CreateConfig
        MapSetKey Config$, "BackgroundColor", MKL$(Config_BackgroundColor)
        MapSetKey Config$, "TextColor", MKL$(Config_TextColor)
        MapSetKey Config$, "CurrentWorkspace", Config_CurrentWorkspace
        MapSetKey Config$, "FileSeperator", Config_FileSeperator
        MapSetKey Config$, "OpenFileSpeed", MKL$(Config_OpenFileSpeed)
End Sub
Sub ParseConfig
        Config_BackgroundColor = CVL(MapGetKey(Config$, "BackgroundColor"))
        Config_TextColor = CVL(MapGetKey(Config$, "TextColor"))
        Config_FileSeperator = MapGetKey(Config$, "FileSeperator")
        Config_OpenFileSpeed = CVL(MapGetKey(Config$, "OpenFileSpeed"))
End Sub
Sub WriteConfigFile
        CreateConfig
        If _FileExists("config.ini") Then Kill "config.ini"
        F = FreeFile
        Open "config.ini" For Output As #F
        Print #F, Config$;
        Close #F
End Sub
'------------------------------

'-------- File Management --------
Sub NewFile
        FileID = UBound(File) + 1: ReDim _Preserve As File File(1 To FileID)
        File(FileID).Name = "Untitled.txt": File(FileID).Path = _StartDir$ + FILE_SEPERATOR + "Untitled.txt"
        File(FileID).Cursors = MKL$(1) + MKL$(1)
        File(FileID).FileID = 0
        File(FileID).TotalLines = 1
        File(FileID).Opened = 1: File(FileID).Saved = 0
        File(FileID).Content = MKL$(GetNewRopePointer)
        File(FileID).ScrollOffset = 1
        File(FileID).HorizontalScrollOffset = 1
End Sub
Sub OpenFile (File$)
        F = 0: If _FileExists(File$) Then F = FreeFile: Open File$ For Input As #F
        FileID = UBound(File) + 1: ReDim _Preserve As File File(1 To FileID)
        File(FileID).Name = GetLastPathName$(File$): File(FileID).Path = File$
        File(FileID).Cursors = MKL$(1) + MKL$(1)
        File(FileID).FileID = F
        File(FileID).TotalLines = 1 - Sgn(F)
        File(FileID).Opened = 1 - Sgn(F): File(FileID).Saved = Sgn(F)
        If F = 0 Then File(FileID).Content = MKL$(GetNewRopePointer) Else File(FileID).Content = ""
        File(FileID).ScrollOffset = 1
        File(FileID).HorizontalScrollOffset = 1
End Sub
Sub OpenFileTasks: Static As Long OpeningFileDialog, OpeningFile_FileNameLabel, OpeningFile_Progress, OpeningFile_Cancel
        If OpeningFileDialog = 0 Then
                OpeningFileDialog = UI_New("Dialog_OpeningFile", UI_TYPE_Dialog, 0, 0, 256, 80, "Opening File", "")
                UI(OpeningFileDialog).Visible = 0
                UI_PARENT = OpeningFileDialog
                OpeningFile_FileNameLabel = UI_New("Dialog_OpeningFile_Label", UI_TYPE_Label, 16, 24, 224, 16, "", "")
                OpeningFile_Progress = UI_New("Dialog_OpeningFile_Progress", UI_TYPE_ProgressBar, 16, 42, 224, 4, "", "")
                OpeningFile_Cancel = UI_New("Dialog_OpeningFile_Cancel", UI_TYPE_Button, 16, 48, 224, 16, "Cancel", "")
                UI_PARENT = 0
        End If

        UI(OpeningFileDialog).Visible = 0
        For FileID = 1 To UBound(File)
                If File(FileID).Opened Then _Continue
                For J = 1 To Config_OpenFileSpeed
                        UI(OpeningFile_FileNameLabel).Label = File(FileID).Name
                        UI(OpeningFileDialog).Visible = -1
                        Line Input #File(FileID).FileID, L$
                        I = GetNewRopePointer
                        Rope(I) = L$
                        UpdateRope I
                        File(FileID).TotalLines = File(FileID).TotalLines + 1
                        If Len(File(FileID).Content) < File(FileID).TotalLines * 4 + 4 Then File(FileID).Content = File(FileID).Content + String$(1024, 0)
                        Mid$(File(FileID).Content, File(FileID).TotalLines * 4 - 3, 4) = MKL$(I)
                        UI(OpeningFile_Progress).Progress = Seek(File(FileID).FileID) / LOF(File(FileID).FileID)
                        If Seek(File(FileID).FileID) > LOF(File(FileID).FileID) Then Close #File(FileID).FileID: File(FileID).Opened = 1: Exit For
                Next J
                If UI(OpeningFile_Cancel).Response Then File(FileID).Opened = 1: UI_Focus = 0
        Next FileID
End Sub
Sub SaveFile (I As Long)
        If I = 0 Then Exit Sub
        SaveFileQueue = SaveFileQueue + MKL$(I)
        File(I).Saved = 0
        File(I).SaveOffset = 0
        File(I).FileID = FreeFile
        Open File(I).Path For Output As #File(I).FileID
End Sub
Sub SaveFileTasks: Static As Long SavingFileDialog, SavingFile_FileNameLabel, SavingFile_Progress
        If SavingFileDialog = 0 Then
                SavingFileDialog = UI_New("Dialog_SavingFile", UI_TYPE_Dialog, 0, 0, 256, 64, "Saving File", "")
                UI(SavingFileDialog).Visible = 0
                UI_PARENT = SavingFileDialog
                SavingFile_FileNameLabel = UI_New("Dialog_SavingFile_Label", UI_TYPE_Label, 16, 24, 224, 16, "", "")
                SavingFile_Progress = UI_New("Dialog_SavingFile_Progress", UI_TYPE_ProgressBar, 16, 48, 224, 4, "", "")
                UI_PARENT = 0
        End If
        UI(SavingFileDialog).Visible = 0
        If Len(SaveFileQueue) = 0 Then Exit Sub
        FileID = CVL(Left$(SaveFileQueue, 4))
        UI(SavingFileDialog).Visible = -1
        UI(SavingFile_FileNameLabel).Label = File(FileID).Name
        If File(FileID).Saved = 0 Then
                For J = 1 To Config_OpenFileSpeed
                        File(FileID).SaveOffset = File(FileID).SaveOffset + 1
                        If File(FileID).SaveOffset > File(FileID).TotalLines Then Exit For
                        RopeI = CVL(Mid$(File(FileID).Content, _SHL(File(FileID).SaveOffset, 2) - 3, 4))
                        If File(FileID).SaveOffset = File(FileID).TotalLines Then Print #File(FileID).FileID, Rope(RopeI); Else Print #File(FileID).FileID, Rope(RopeI) + Config_FileSeperator;
                Next J
                UI(SavingFile_Progress).Progress = File(FileID).SaveOffset / File(FileID).TotalLines
        End If
        If File(FileID).SaveOffset < File(FileID).TotalLines Then Exit Sub
        File(FileID).Saved = 1
        SaveFileQueue = Mid$(SaveFileQueue, 5)
        Close #File(FileID).FileID
End Sub
Sub CloseFile (F As Long)
        EmptyRopePoints = EmptyRopePoints + Left$(File(F).Content, _SHL(File(F).TotalLines, 2))
        For I = F + 1 To UBound(File): Swap File(I), File(I - 1): Next I
        If I > 2 Then ReDim _Preserve File(1 To I - 2) As File Else ReDim File(0) As File
End Sub
Function GetNewRopePointer&
        If Len(EmptyRopePoints) < 4 Then
                I = UBound(Rope) + 1
                ReDim _Preserve Rope(1 To I) As String
                GetNewRopePointer& = I
        Else
                I = CVL(Left$(EmptyRopePoints, 4))
                Rope(I) = ""
                EmptyRopePoints = Mid$(EmptyRopePoints, 5)
                GetNewRopePointer& = I
        End If
End Function
'---------------------------------

'-------- Workspace Management --------
Sub OpenWorkspace (Path$)
        __F = FreeFile
        Open Path$ For Binary As #__F
        WorkspaceMap$ = String$(LOF(__F), 0)
        Get #__F, , WorkspaceMap$
        Close #__F
        If Len(WorkspaceMap$) < 5 Then Exit Sub
        For I = 1 To CVI(MapGetKey(WorkspaceMap$, "TotalFiles"))
                FileMap$ = MapGetKey(WorkspaceMap$, MKL$(I))
                OpenFile MapGetKey(FileMap$, "Path")
                FileID = UBound(File)
                File(FileID).Cursors = MapGetKey(FileMap$, "Cursors")
                File(FileID).ScrollOffset = CVL(MapGetKey(FileMap$, "ScrollOffset"))
        Next I
End Sub
Sub SaveWorkspace (Path$)
        WorkspaceMap$ = MapNew
        MapSetKey WorkspaceMap$, "TotalFiles", MKI$(UBound(File))
        For I = 1 To UBound(File)
                FileMap$ = MapNew
                MapSetKey FileMap$, "Path", File(I).Path
                MapSetKey FileMap$, "Cursors", File(I).Cursors
                MapSetKey FileMap$, "ScrollOffset", MKL$(File(I).ScrollOffset)
                MapSetKey WorkspaceMap$, MKL$(I), FileMap$
        Next I
        __F = FreeFile
        Open Path$ For Output As #__F
        Print #__F, WorkspaceMap$;
        Close #__F
End Sub
'--------------------------------------

'----------------------------------------------------------------Libraries----------------------------------------------------------------
Sub WaitForMouseButtonRelease
        While _MouseInput Or _MouseButton(1): Wend
End Sub
Sub PrintString (X As Integer, Y As Integer, S As String)
        _PrintString (X, Y), S
End Sub
Function GetLastPathName$ (Path$)
        T~& = _InStrRev(Path$, FILE_SEPERATOR) + 1
        GetLastPathName$ = Mid$(Path$, T~&)
End Function
Function GetKeyDown (__Key As _Unsigned Long)
        If __Key = 0 Then Exit Function
        GetKeyDown = _KeyDown(__Key)
End Function
Function ceil& (t#)
        t& = Int(t#)
        ceil& = t& + Sgn(t# - t&)
End Function

Function PathBack$ (Path$)
        If Right$(Path$, 1) = FILE_SEPERATOR Then P$ = Left$(Path$, Len(Path$) - 1) Else P$ = Path$
        PathBack$ = Left$(Path$, _InStrRev(P$, FILE_SEPERATOR))
End Function
Function GetDirList$ (CurrentDirectory$)
        If CurrentDirectory$ = "" Or CurrentDirectory$ = "\" Then
                $If WIN Then
                        O$ = ListStringNew
                        For I = 65 To 90
                                If _DirExists(Chr$(I) + ":\") Then ListStringAdd O$, Chr$(I) + ":\"
                        Next I
                        GetDirList$ = O$
                        CurrentDirectory$ = ""
                        Exit Function
                $Else
                                CurrentDirectory$ = "/"
                $End If
        End If
        If Len(CurrentDirectory$) > 1 And Right$(CurrentDirectory$, 1) <> FILE_SEPERATOR Then CurrentDirectory$ = CurrentDirectory$ + FILE_SEPERATOR
        O$ = ListStringNew: B = 0
        T = _ShellHide("dir /b /o:n /a:d " + Chr$(34) + CurrentDirectory$ + Chr$(34) + " > tmp.txt & echo.>> tmp.txt & dir /b /o:n /a:-d " + Chr$(34) + CurrentDirectory$ + Chr$(34) + " >> tmp.txt")
        __F = FreeFile: Open "tmp.txt" For Input As #__F: Do While Seek(__F) < LOF(__F)
                Line Input #__F, L$: If L$ = " " Then B = 1: _Continue
                If Left$(CurrentDirectory$, Len(_CWD$)) = _CWD$ And L$ = "tmp.txt" Then _Continue
                If B Then ListStringAdd O$, L$ Else ListStringAdd O$, L$ + FILE_SEPERATOR
        Loop: Close #__F: Kill "tmp.txt"
        GetDirList$ = O$
End Function
Sub SetTitle (T$)
        Static oldTitle As String
        If oldTitle <> T$ Then _Title T$
        oldTitle = T$
End Sub

Function UI_New (__Name As String, __Type As _Unsigned _Byte, __X As Integer, __Y As Integer, __W As Integer, __H As Integer, __Label As String, __Content As String)
        I = UBound(UI) + 1
        ReDim _Preserve UI(1 To I) As UI
        UI(I).Parent = UI_PARENT
        UI(I).Name = __Name
        UI(I).Type = __Type
        UI(I).X = __X
        UI(I).Y = __Y
        UI(I).W = __W
        UI(I).H = __H
        UI(I).Label = __Label
        If __Type = UI_TYPE_Dialog Then UI(I).Property = UI_PROP_Center Else UI(I).Property = 0
        If InStr(UI(I).Label, "\") Then
                UI(I).ParsedLabel = ""
                UI(I).ParsedLabelUnderlined = ""
                For J = 1 To Len(UI(I).Label)
                        B$1 = Mid$(UI(I).Label, J, 1)
                        If B$1 = "\" Then
                                UI(I).ParsedLabelUnderlined = UI(I).ParsedLabelUnderlined + "_"
                        Else
                                UI(I).ParsedLabelUnderlined = UI(I).ParsedLabelUnderlined + " "
                                UI(I).ParsedLabel = UI(I).ParsedLabel + B$1
                        End If
                Next J
        Else
                UI(I).ParsedLabel = UI(I).Label
        End If
        UI(I).Content = __Content
        UI(I).ParsedContent = ""
        UI(I).Visible = -1
        UI(I).BG = _BackgroundColor
        UI(I).FG = _DefaultColor
        UI(I).OnKey = UI_KEY_Focus
        UI(I).Key0_1 = 0
        UI(I).Key0_2 = 0
        UI(I).Key1_1 = 0
        UI(I).Key1_2 = 0
        UI(I).State = 0
        UI(I).Scroll = 1
        UI(I).MaxScroll = 0
        UI(I).TotalStates = 0
        If Len(__Content) Then If Asc(__Content) = 1 Then UI_ParseContent
        UI_New = I
End Function
Sub UI_Set_FG (Colour As Long)
        UI(UBound(UI)).FG = Colour
End Sub
Sub UI_Set_BG (Colour As Long)
        UI(UBound(UI)).BG = Colour
End Sub
Sub UI_Set_TotalStates (TotalStates As _Unsigned _Byte)
        UI(UBound(UI)).TotalStates = TotalStates
End Sub
Sub UI_Set_KeyMaps (A As Long, B As Long, C As Long, D As Long)
        I = UBound(UI)
        UI(I).Key0_1 = A: UI(I).Key0_2 = B
        UI(I).Key1_1 = C: UI(I).Key1_2 = D
End Sub
Sub UI_ParseContent
        I = UBound(UI)
        L~& = ListStringLength(UI(I).Content)
        Select Case UI(I).Type
                Case UI_TYPE_MenuButton: UI(I).MenuDialogWidth = 0
                        tmpDialogWidth = 0
                        UI(I).ParsedContent = ListStringNew
                        UI(I).ParsedContentUnderlined = ListStringNew
                        For K~& = 1 To L~&
                                K$ = ListStringGet(UI(I).Content, K~&)
                                If InStr(K$, "\") Then
                                        Parsed$ = ""
                                        ParsedUnderlined$ = ""
                                        For J = 1 To Len(K$)
                                                B$1 = Mid$(K$, J, 1)
                                                If B$1 = "\" Then
                                                        ParsedUnderlined$ = ParsedUnderlined$ + "_"
                                                Else
                                                        ParsedUnderlined$ = ParsedUnderlined$ + " "
                                                        Parsed$ = Parsed$ + B$1
                                                End If
                                        Next J
                                Else
                                        Parsed$ = K$
                                        ParsedUnderlined$ = ""
                                End If
                                T~& = InStr(Parsed$, "|")
                                If T~& Then
                                        UI(I).MenuDialogWidth = Max(UI(I).MenuDialogWidth, T~& - 1)
                                        tmpDialogWidth = Max(tmpDialogWidth, Len(Parsed$) - T~&)
                                Else
                                        UI(I).MenuDialogWidth = Max(UI(I).MenuDialogWidth, Len(Parsed$))
                                End If
                                ListStringAdd UI(I).ParsedContent, Space$(1) + Parsed$
                                ListStringAdd UI(I).ParsedContentUnderlined, Space$(1) + ParsedUnderlined$
                        Next K~&
                        UI(I).MenuDialogWidth = UI(I).MenuDialogWidth + 3
                        For K~& = 1 To L~&
                                Parsed$ = ListStringGet(UI(I).ParsedContent, K~&)
                                T~& = InStr(Parsed$, "|")
                                If T~& = 0 Then _Continue
                                Parsed$ = Left$(Parsed$, T~& - 1) + Space$(Max(0, UI(I).MenuDialogWidth - T~& + 1)) + Mid$(Parsed$, T~& + 1)
                                ListStringEdit UI(I).ParsedContent, Parsed$, K~&
                        Next K~&
                        UI(I).MenuDialogWidth = (UI(I).MenuDialogWidth + tmpDialogWidth + 2) * _FontWidth
                        UI(I).MenuDialogHeight = (K~& + 1) * _FontHeight
                Case UI_TYPE_ToggleButton
                        UI(I).TotalStates = L~&
        End Select
End Sub
Sub UI_Draw
        Dim As UI __UI
        Dim As _Unsigned Long __FG, __BG
        __FG = _DefaultColor
        __BG = _BackgroundColor
        If UI_Focus > 0 Then If UI(UI_Focus).Visible = 0 Then UI_Focus = 0
        For I = 1 To UBound(UI)
                UI(I).Response = 0
                If UI(I).Parent = 0 Then
                        Parent_Width = _Width
                        Parent_Height = _Height
                        Parent_X = 0
                        Parent_Y = 0
                Else
                        Parent_Width = UI(UI(I).Parent).__W
                        Parent_Height = UI(UI(I).Parent).__H
                        Parent_X = UI(UI(I).Parent).__X
                        Parent_Y = UI(UI(I).Parent).__Y
                End If
                __UI = UI(I)
                If __UI.Property = UI_PROP_Center Then
                        If __UI.W < 0 Then __UI.__W = Parent_Width + __UI.W Else __UI.__W = __UI.W
                        If __UI.H < 0 Then __UI.__H = Parent_Height + __UI.H Else __UI.__H = __UI.H
                        __UI.__X = (Parent_Width - __UI.__W) / 2
                        __UI.__Y = (Parent_Height - __UI.__H) / 2
                Else
                        __UI.__X = ClampCycleDifference(0, __UI.X, Parent_Width)
                        __UI.__Y = ClampCycleDifference(0, __UI.Y, Parent_Height)
                        If __UI.W < 0 Then __UI.__W = Parent_Width + __UI.W - __UI.__X Else __UI.__W = __UI.W
                        If __UI.H < 0 Then __UI.__H = Parent_Height + __UI.H - __UI.__Y Else __UI.__H = __UI.H
                End If
                __UI.__X = __UI.__X + Parent_X
                __UI.__Y = __UI.__Y + Parent_Y
                If UI(I).Parent Then
                        If UI(UI(I).Parent).Visible = 0 Then _Continue
                End If
                UI_Focus_Parent = 0
                Select Case __UI.Type And __UI.Visible
                        Case UI_TYPE_Label
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, B
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) Then
                                        Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                        If _MouseButton(1) Then __UI.Response = -1: WaitForMouseButtonRelease
                                End If
                                PrintString __UI.__X, __UI.__Y + (__UI.__H - _FontHeight) / 2, __UI.Label

                        Case UI_TYPE_Button
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), _RGB32(255, 127), B
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) Then
                                        Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                        If _MouseButton(1) Then UI_Focus = I: __UI.Response = -1: WaitForMouseButtonRelease
                                End If
                                PrintString __UI.__X + (__UI.W - _FontWidth * Len(__UI.Label)) / 2, __UI.__Y + (__UI.__H - _FontHeight) / 2, __UI.Label

                        Case UI_TYPE_Dialog
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), _RGB32(0, 191), BF
                                Line (__UI.__X + _FontWidth / 2, __UI.__Y + _FontHeight / 2)-(__UI.__X + 1.5 * _FontWidth, __UI.__Y + _FontHeight / 2), __UI.FG
                                Line (__UI.__X + __UI.__W - _FontWidth / 2, __UI.__Y + _FontHeight / 2)-(__UI.__X + _FontWidth / 2 + _FontWidth * (Len(__UI.Label) + 2), __UI.__Y + _FontHeight / 2), __UI.FG
                                Line (__UI.__X + _FontWidth / 2, __UI.__Y + _FontHeight / 2)-(__UI.__X + _FontWidth / 2, __UI.__Y + __UI.__H - _FontHeight / 2), __UI.FG
                                Line (__UI.__X + _FontWidth / 2, __UI.__Y + __UI.__H - _FontHeight / 2)-(__UI.__X + __UI.__W - _FontWidth / 2, __UI.__Y + __UI.__H - _FontHeight / 2), __UI.FG
                                Line (__UI.__X + __UI.__W - _FontWidth / 2, __UI.__Y + __UI.__H - _FontHeight / 2)-(__UI.__X + __UI.__W - _FontWidth / 2, __UI.__Y + _FontHeight / 2), __UI.FG
                                PrintString __UI.__X + 2 * _FontWidth, __UI.__Y, __UI.Label
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) = 0 And _MouseButton(1) Then UI_Focus = __UI.Parent: __UI.Response = -1: WaitForMouseButtonRelease
                                If KeyHit = 27 Then UI_Focus = __UI.Parent: __UI.Response = -1

                        Case UI_TYPE_ProgressBar
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W * __UI.Progress, __UI.__Y + __UI.__H), __UI.FG, BF

                        Case UI_TYPE_MenuButton
                                If UI_Focus = I Then Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                PrintString __UI.__X, __UI.__Y, __UI.ParsedLabel
                                PrintString __UI.__X, __UI.__Y, __UI.ParsedLabelUnderlined
                                If UI_Focus = __UI.Parent And MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = I: WaitForMouseButtonRelease
                                If UI_Focus = I Then
                                        Line (__UI.__X, __UI.__Y + __UI.__H)-(__UI.__X + __UI.MenuDialogWidth, __UI.__Y + __UI.__H + __UI.MenuDialogHeight), _RGB32(0, 223), BF
                                        Line (__UI.__X + _FontWidth / 2, __UI.__Y + __UI.__H + _FontHeight / 2)-(__UI.__X + __UI.MenuDialogWidth - _FontWidth / 2, __UI.__Y + __UI.__H + __UI.MenuDialogHeight - _FontHeight / 2), _RGB32(255), B
                                        For J = 1 To ListStringLength(__UI.ParsedContent)
                                                X1 = __UI.__X + _FontWidth
                                                Y1 = __UI.__Y + (1 + J) * _FontHeight
                                                X2 = __UI.__X + __UI.MenuDialogWidth - _FontWidth
                                                Y2 = __UI.__Y + (2 + J) * _FontHeight - 1
                                                P$ = ListStringGet(__UI.ParsedContent, J)
                                                U$ = ListStringGet(__UI.ParsedContentUnderlined, J)
                                                If MouseInBox(X1, Y1, X2, Y2) Then
                                                        Line (X1, Y1)-(X2, Y2), __UI.BG, BF
                                                        If _MouseButton(1) Then
                                                                WaitForMouseButtonRelease
                                                                UI_Focus = __UI.Parent
                                                                __UI.Response = J
                                                        End If
                                                End If
                                                If InRange(32, KeyHit, 122) And InStr(U$, "_") Then
                                                        If LCase$(Chr$(KeyHit)) = LCase$(Mid$(P$, InStr(U$, "_"), 1)) Then
                                                                UI_Focus = __UI.Parent
                                                                __UI.Response = J
                                                        End If
                                                End If
                                                PrintString X1, Y1, P$
                                                PrintString X1, Y1, U$
                                        Next J
                                        If KeyHit = 27 Then UI_Focus = __UI.Parent: __UI.Response = -1
                                        If MouseInBox(__UI.__X, __UI.__Y + __UI.__H, __UI.__X + __UI.MenuDialogWidth, __UI.__Y + __UI.__H + __UI.MenuDialogHeight) = 0 And _MouseButton(1) Then UI_Focus = __UI.Parent
                                End If

                        Case UI_TYPE_Image
                                If UI_Focus = __UI.Parent And MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = I: WaitForMouseButtonRelease
                                _PutImage (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.Image

                        Case UI_TYPE_ScrollBar
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                If __UI.MaxScroll < 2 Then __Y = 0 Else __Y = (__UI.__H - 16) * (__UI.Scroll - 1) / (__UI.MaxScroll - 1)
                                Line (__UI.__X + 1, __UI.__Y + __Y)-(__UI.__X + __UI.__W - 1, __UI.__Y + __Y + 16), _RGB32(255), BF
                                If UI_Focus = __UI.Parent And MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then __UI.Scroll = Clamp(1, (_MouseY - __UI.__Y - 8) * __UI.MaxScroll / (__UI.__H - 16), __UI.MaxScroll): __UI.Response = -1

                        Case UI_TYPE_ToggleButton
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) Then
                                        Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                        If _MouseButton(1) Then
                                                __UI.State = ClampCycle(0, __UI.State + 1, __UI.TotalStates - 1)
                                                __UI.Response = __UI.State + 1
                                                WaitForMouseButtonRelease
                                        End If
                                End If
                                PrintString __UI.__X, __UI.__Y, ListStringGet(__UI.Content, __UI.State + 1)

                        Case UI_TYPE_ListView
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = __UI.Parent
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.FG, B
                                If UI_Focus = __UI.Parent Then __UI.Scroll = Clamp(1, __UI.Scroll + UI_MouseWheel, ListStringLength(__UI.ParsedContent))
                                K = 0
                                For J = Max(1, __UI.Scroll) To Min(__UI.Scroll + __UI.__H / _FontHeight - 1, ListStringLength(__UI.ParsedContent))
                                        X1 = __UI.__X + _FontWidth
                                        Y1 = __UI.__Y + K
                                        X2 = __UI.__X + __UI.__W - _FontWidth
                                        Y2 = __UI.__Y + K + _FontHeight - 1
                                        If MouseInBox(X1, Y1, X2, Y2) Then
                                                Line (X1, Y1)-(X2, Y2), __UI.BG, BF
                                                If _MouseButton(1) Then
                                                        __UI.Selected = J
                                                        __UI.Response = J
                                                        WaitForMouseButtonRelease
                                                End If
                                        End If
                                        If __UI.Selected = J Then Line (X1, Y1)-(X2, Y2), __UI.BG, BF
                                        L$ = ListStringGet(__UI.ParsedContent, J)
                                        If Len(L$) > _SHR(X2 - X1, 3) Then PrintString X1, Y1, Left$(L$, _SHR(X2 - X1, 3)) Else PrintString X1, Y1, L$
                                        K = K + _FontHeight
                                Next J

                        Case UI_TYPE_TextView
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), _RGB32(255), B
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = I: UI_Focus_Parent = __UI.Parent: WaitForMouseButtonRelease
                                If UI_Focus = I Then
                                        Select Case KeyHit
                                                Case 8: __UI.Content = Left$(__UI.Content, Len(__UI.Content) - 1)
                                                Case 9, 27: UI_Focus = __UI.Parent
                                                Case 32 To 122: __UI.Content = __UI.Content + Chr$(KeyHit)
                                        End Select
                                        __UI.Response = KeyHit
                                End If
                                PrintString __UI.__X, __UI.__Y, __UI.Content + "_"

                        Case UI_TYPE_Frame
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, B

                End Select
                If (GetKeyDown(__UI.Key0_1) Or GetKeyDown(__UI.Key1_1)) And (GetKeyDown(__UI.Key0_2) Or GetKeyDown(__UI.Key1_2)) Then
                        Select Case __UI.OnKey
                                Case UI_KEY_Visible: __UI.Visible = -1
                                Case UI_KEY_Focus: UI_Focus = I
                                Case UI_KEY_Visible_And_Focus: __UI.Visible = -1: UI_Focus = I
                        End Select
                End If
                UI(I) = __UI
        Next I
        Color __FG, __BG
End Sub
Function MouseInBox (X1, Y1, X2, Y2)
        MouseInBox = X1 <= _MouseX And _MouseX <= X2 And Y1 <= _MouseY And _MouseY <= Y2
End Function
Function Max (A, B)
        Max = -A * (A > B) - B * (A <= B)
End Function
Function Min (A, B)
        Min = -A * (A < B) - B * (A >= B)
End Function
Function Clamp (A, B, C)
        Clamp = B - (A - B) * (B < A) - (C - B) * (C < B)
End Function
Function ClampCycle (A, B, C)
        ClampCycle = B - (C - B) * (B < A) - (A - B) * (C < B)
End Function
Function ClampCycleDifference (A, B, C)
        ClampCycleDifference = B - C * (B < A) - A * (C < B)
End Function
Function InRange (A, B, C)
        InRange = (A <= B) And (B <= C)
End Function
Function CeilDiv (A As Long, B As Integer)
        Static D As Long
        D = A \ B
        CeilDiv = D + (A - D * B)
End Function
Function CeilDiv2 (A As Long, B As _Byte)
        Static D As Long
        D = _SHR(A, B)
        CeilDiv2 = D + (A - _SHL(D, B))
End Function
Function ListStringNew$
        ListStringNew$ = Chr$(1) + MKL$(0)
End Function
Function ListStringFromString$ (ARRAY$)
        LIST$ = Chr$(1) + MKL$(0)
        __LISTLENGTH~& = 0
        For I~& = 1 To Len(ARRAY$)
                BYTE~%% = Asc(ARRAY$, I~&)
                Select Case BYTE~%%
                        Case 91: NEST~& = NEST~& + 1
                                V$ = V$ + Chr$(BYTE~%%)
                        Case 93: NEST~& = NEST~& - 1
                                V$ = V$ + Chr$(BYTE~%%)
                        Case 44: If NEST~& = 0 Then
                                        __LISTLENGTH~& = __LISTLENGTH~& + 1
                                        LIST$ = LIST$ + MKI$(Len(V$)) + V$
                                        V$ = ""
                                Else
                                        V$ = V$ + Chr$(BYTE~%%)
                                End If
                        Case Else:
                                V$ = V$ + Chr$(BYTE~%%)
                End Select
        Next I~&
        If Len(V$) Then
                LIST$ = LIST$ + MKI$(Len(V$)) + V$
                __LISTLENGTH~& = __LISTLENGTH~& + 1
                V$ = ""
        End If
        Mid$(LIST$, 2, 4) = MKL$(__LISTLENGTH~&)
        ListStringFromString$ = LIST$
        LIST$ = ""
End Function
Function ListStringFromArray$ (ARRAY() As String, START_INDEX~&, END_INDEX~&)
        K~& = 0
        For I~& = START_INDEX~& To END_INDEX~&
                K~& = K~& + 2 + Len(ARRAY(I~&))
        Next I~&
        LIST$ = Chr$(1) + MKL$(END_INDEX~& - START_INDEX~& + 1) + String$(K~&, 0)
        K~& = 6
        L~% = 0
        For I~& = START_INDEX~& To END_INDEX~&
                L~% = Len(ARRAY(I~&))
                Mid$(LIST$, K~&, 2 + L~%) = MKI$(L~%) + ARRAY(I~&)
                K~& = K~& + L~% + 2
        Next I~&
        ListStringFromArray$ = LIST$
        LIST$ = ""
End Function
Function ListStringPrint$ (LIST$)
        If Len(LIST$) < 5 Then Exit Function
        If Asc(LIST$) <> 1 Then Exit Function
        O = 6: T_OFFSET = 2
        T$ = String$(Len(LIST$) - 4, 0)
        Asc(T$) = 91 '[
        For I = 1 To CVL(Mid$(LIST$, 2, 4)) - 1
                L = CVI(Mid$(LIST$, O, 2))
                Mid$(T$, T_OFFSET, L + 1) = Mid$(LIST$, O + 2, L) + ","
                T_OFFSET = T_OFFSET + L + 1
                O = O + L + 2
        Next I
        L = CVI(Mid$(LIST$, O, 2))
        Mid$(T$, T_OFFSET, L + 1) = Mid$(LIST$, O + 2, L) + "]"
        ListStringPrint$ = Left$(T$, T_OFFSET + L)
End Function
Function ListStringLength~& (LIST$)
        If Len(LIST$) < 5 Then Exit Function
        If Asc(LIST$) <> 1 Then Exit Function
        ListStringLength~& = CVL(Mid$(LIST$, 2, 4))
End Function
Sub ListStringAdd (LIST$, ITEM$)
        If Len(LIST$) < 5 Then Exit Sub
        If Asc(LIST$) <> 1 Then Exit Sub
        LIST$ = Chr$(1) + MKL$(CVL(Mid$(LIST$, 2, 4)) + 1) + Mid$(LIST$, 6) + MKI$(Len(ITEM$)) + ITEM$
End Sub
Function ListStringGet$ (LIST$, POSITION As _Unsigned Long)
        If Len(LIST$) < 5 Then Exit Function
        If Asc(LIST$) <> 1 Then Exit Function
        If CVL(Mid$(LIST$, 2, 4)) < POSITION - 1 Then Exit Function
        O = 6
        For I = 1 To POSITION - 1
                L = CVI(Mid$(LIST$, O, 2))
                O = O + L + 2
        Next I
        ListStringGet$ = Mid$(LIST$, O + 2, CVI(Mid$(LIST$, O, 2)))
End Function
Sub ListStringInsert (LIST$, ITEM$, POSITION As _Unsigned Long)
        If Len(LIST$) < 5 Then Exit Sub
        If Asc(LIST$) <> 1 Then Exit Sub
        If CVL(Mid$(LIST$, 2, 4)) < POSITION - 1 Then Exit Sub
        O = 6
        For I = 1 To POSITION - 1
                L = CVI(Mid$(LIST$, O, 2))
                O = O + L + 2
        Next I
        LIST$ = Chr$(1) + MKL$(CVL(Mid$(LIST$, 2, 4)) + 1) + Mid$(LIST$, 6, O - 6) + MKI$(Len(ITEM$)) + ITEM$ + Mid$(LIST$, O)
End Sub
Sub ListStringDelete (LIST$, POSITION As _Unsigned Long)
        If Len(LIST$) < 5 Then Exit Sub
        If Asc(LIST$) <> 1 Then Exit Sub
        If CVL(Mid$(LIST$, 2, 4)) < POSITION - 1 Then Exit Sub
        O = 6
        For I = 1 To POSITION - 1
                L = CVI(Mid$(LIST$, O, 2))
                O = O + L + 2
        Next I
        LIST$ = Chr$(1) + MKL$(CVL(Mid$(LIST$, 2, 4)) - 1) + Mid$(LIST$, 6, O - 6) + Mid$(LIST$, O + CVI(Mid$(LIST$, O, 2)) + 2)
End Sub
Function ListStringSearch~& (LIST$, ITEM$)
        If Len(LIST$) < 5 Then Exit Function
        If Asc(LIST$) <> 1 Then Exit Function
        O = 6
        For I = 1 To CVL(Mid$(LIST$, 2, 4))
                L = CVI(Mid$(LIST$, O, 2))
                If ITEM$ = Mid$(LIST$, O + 2, L) Then ListStringSearch~& = I: Exit Function
                O = O + L + 2
        Next I
        ListStringSearch~& = 0
End Function
Function ListStringSearchI~& (LIST$, ITEM$)
        If Len(LIST$) < 5 Then Exit Function
        If Asc(LIST$) <> 1 Then Exit Function
        O = 6
        For I = 1 To CVL(Mid$(LIST$, 2, 4))
                L = CVI(Mid$(LIST$, O, 2))
                If _StriCmp(ITEM$, Mid$(LIST$, O + 2, L)) = 0 Then ListStringSearchI~& = I: Exit Function
                O = O + L + 2
        Next I
        ListStringSearchI~& = 0
End Function
Sub ListStringEdit (LIST$, ITEM$, POSITION As _Unsigned Long)
        If Len(LIST$) < 5 Then Exit Sub
        If Asc(LIST$) <> 1 Then Exit Sub
        If CVL(Mid$(LIST$, 2, 4)) < POSITION - 1 Then Exit Sub
        O = 6
        For I = 1 To POSITION - 1
                L = CVI(Mid$(LIST$, O, 2))
                O = O + L + 2
        Next I
        LIST$ = Left$(LIST$, O - 1) + MKI$(Len(ITEM$)) + ITEM$ + Mid$(LIST$, O + CVI(Mid$(LIST$, O, 2)) + 2)
End Sub
Function ListStringAppend$ (LIST1$, LIST2$)
        If Len(LIST1$) < 5 Then Exit Function
        If Len(LIST2$) < 5 Then Exit Function
        If Asc(LIST1$) <> 1 Then Exit Function
        If Asc(LIST2$) <> 1 Then Exit Function
        ListStringAppend$ = Chr$(1) + MKL$(CVL(Mid$(LIST1$, 2, 4)) + CVL(Mid$(LIST2$, 2, 4))) + Mid$(LIST1$, 6) + Mid$(LIST2$, 6)
End Function
Function MapNew$
        MapNew$ = Chr$(8) + MKL$(0)
End Function
Sub MapSetKey (__MAP$, __KEY$, __VALUE$)
        If Len(__MAP$) < 5 Then Exit Sub
        If Asc(__MAP$, 1) <> 8 Then Exit Sub
        Dim __LENGTH~&, __OFFSET~&
        __LENGTH~& = CVL(Mid$(__MAP$, 2, 4))
        __OFFSET~& = 6
        Dim __LEN_KEY~&, __LEN_VALUE~&
        For __I~& = 1 To __LENGTH~& 'check if exists
                __LEN_KEY~& = CVL(Mid$(__MAP$, __OFFSET~&, 4))
                __LEN_VALUE~& = CVL(Mid$(__MAP$, __OFFSET~& + 4, 4))
                __K$ = Mid$(__MAP$, __OFFSET~& + 8, __LEN_KEY~&)
                If __KEY$ = __K$ Then __BOOL` = -1: Exit For
                __V$ = Mid$(__MAP$, __OFFSET~& + 8 + __LEN_KEY~&, __LEN_VALUE~&)
                __OFFSET~& = __OFFSET~& + 8 + __LEN_KEY~& + __LEN_VALUE~&
        Next __I~&
        If __BOOL` Then
                __MAP$ = Left$(__MAP$, __OFFSET~& - 1) + MKL$(Len(__KEY$)) + MKL$(Len(__VALUE$)) + __KEY$ + __VALUE$ + Mid$(__MAP$, __OFFSET~& + 8 + __LEN_KEY~& + __LEN_VALUE~&)
        Else
                __MAP$ = Chr$(8) + MKL$(__LENGTH~& + 1) + Mid$(__MAP$, 6, __OFFSET~& - 6) + MKL$(Len(__KEY$)) + MKL$(Len(__VALUE$)) + __KEY$ + __VALUE$
        End If
End Sub
Function MapPrint$ (__MAP$)
        If Len(__MAP$) < 5 Then Exit Function
        If Asc(__MAP$, 1) <> 8 Then Exit Function
        Dim __LENGTH~&, __OFFSET~&
        __LENGTH~& = CVL(Mid$(__MAP$, 2, 4))
        __OFFSET~& = 6
        Dim __LEN_KEY~&, __LEN_VALUE~&
        __PRINT$ = ""
        For __I~& = 1 To __LENGTH~& 'check if exists
                __LEN_KEY~& = CVL(Mid$(__MAP$, __OFFSET~&, 4))
                __LEN_VALUE~& = CVL(Mid$(__MAP$, __OFFSET~& + 4, 4))
                __K$ = Mid$(__MAP$, __OFFSET~& + 8, __LEN_KEY~&)
                __V$ = Mid$(__MAP$, __OFFSET~& + 8 + __LEN_KEY~&, __LEN_VALUE~&)
                __OFFSET~& = __OFFSET~& + 8 + __LEN_KEY~& + __LEN_VALUE~&
                __PRINT$ = __PRINT$ + __K$ + ":" + __V$
                If __I~& < __LENGTH~& Then __PRINT$ = __PRINT$ + ","
        Next __I~&
        MapPrint$ = "{" + __PRINT$ + "}"
End Function
Function MapGetKey$ (__MAP$, __KEY$)
        If Len(__MAP$) < 5 Then Exit Function
        If Asc(__MAP$, 1) <> 8 Then Exit Function
        Dim __LENGTH~&, __OFFSET~&
        __LENGTH~& = CVL(Mid$(__MAP$, 2, 4))
        __OFFSET~& = 6
        Dim __LEN_KEY~&, __LEN_VALUE~&
        For __I~& = 1 To __LENGTH~& 'check if exists
                __LEN_KEY~& = CVL(Mid$(__MAP$, __OFFSET~&, 4))
                __LEN_VALUE~& = CVL(Mid$(__MAP$, __OFFSET~& + 4, 4))
                __K$ = Mid$(__MAP$, __OFFSET~& + 8, __LEN_KEY~&)
                __V$ = Mid$(__MAP$, __OFFSET~& + 8 + __LEN_KEY~&, __LEN_VALUE~&)
                If __KEY$ = __K$ Then MapGetKey$ = __V$: Exit Function
                __OFFSET~& = __OFFSET~& + 8 + __LEN_KEY~& + __LEN_VALUE~&
        Next __I~&
End Function
Sub MapDeleteKey (__MAP$, __KEY$)
        If Len(__MAP$) < 5 Then Exit Sub
        If Asc(__MAP$, 1) <> 8 Then Exit Sub
        Dim __LENGTH~&, __OFFSET~&
        __LENGTH~& = CVL(Mid$(__MAP$, 2, 4))
        __OFFSET~& = 6
        Dim __LEN_KEY~&, __LEN_VALUE~&
        For __I~& = 1 To __LENGTH~& 'check if exists
                __LEN_KEY~& = CVL(Mid$(__MAP$, __OFFSET~&, 4))
                __LEN_VALUE~& = CVL(Mid$(__MAP$, __OFFSET~& + 4, 4))
                __K$ = Mid$(__MAP$, __OFFSET~& + 8, __LEN_KEY~&)
                If __KEY$ = __K$ Then __BOOL` = -1: Exit For
                __V$ = Mid$(__MAP$, __OFFSET~& + 8 + __LEN_KEY~&, __LEN_VALUE~&)
                __OFFSET~& = __OFFSET~& + 8 + __LEN_KEY~& + __LEN_VALUE~&
        Next __I~&
        If __BOOL` Then __MAP$ = Chr$(8) + MKL$(__LENGTH~& - 1) + Mid$(__MAP$, 6, __OFFSET~& - 6) + Mid$(__MAP$, __OFFSET~& + 8 + __LEN_KEY~& + __LEN_VALUE~&)
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
