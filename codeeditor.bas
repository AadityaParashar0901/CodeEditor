'$Dynamic
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
        As _Unsigned Long TotalLines
        As Long ScrollOffset, FileID, SaveOffset
        As _Unsigned _Byte Opened, Saved
End Type
Dim Shared File(0) As File, CurrentFile As Long
Dim Shared Rope(0) As String, EmptyRopePoints As String

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
'--------------------

Dim Shared As String SaveFileQueue

Dim Shared Config$, Config_BackgroundColor As Long, Config_TextColor As Long, Config_CurrentWorkspace As String, Config_FileSeperator As String
NewConfig
CreateConfig
ReadConfigFile

'-------- Menu Bar --------
UI_MenuButton_File = UI_New("UI_MenuButton_File", UI_TYPE_MenuButton, 0, 0, 48, 16, " File ", ListStringFromString("\New File|Ctrl+N,\Open File|Ctrl+O,\Save File|Ctrl+S,Save \As File,\Close File|Ctrl+W,E\xit|Ctrl+Q"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 70, 100307, 102
UI_MenuButton_Workspace = UI_New("UI_MenuButton_Workspace", UI_TYPE_MenuButton, 48, 0, 88, 16, " Workspace ", ListStringFromString("\Open Workspace,\Save Workspace,\Close Workspace"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 87, 100307, 119
UI_MenuButton_View = UI_New("UI_MenuButton_View", UI_TYPE_MenuButton, 136, 0, 48, 16, " View ", ListStringFromString("Go to File,Go to Line|Ctrl+G,Toggle Side Pane|Alt+T"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 86, 100307, 118
UI_MenuButton_Cursors = UI_New("UI_MenuButton_Cursors", UI_TYPE_MenuButton, 184, 0, 72, 16, " Cursors ", ListStringFromString("Go to Cursor|Alt+C"))
UI_Set_BG _RGB32(0, 191, 0)
UI_Set_KeyMaps 100308, 67, 100307, 99
'--------------------------
'-------- Status Bar --------
UI_ToggleButton_LineSeperator = UI_New("UI_ToggleButton_LineSeperator", UI_TYPE_ToggleButton, 16, -16, 48, 16, "", ListStringFromString("[CRLF],[  LF],[CR  ]"))
Select Case Config_FileSeperator
        Case Chr$(13) + Chr$(10): UI(UI_ToggleButton_LineSeperator).State = 0
        Case Chr$(10): UI(UI_ToggleButton_LineSeperator).State = 1
        Case Chr$(13): UI(UI_ToggleButton_LineSeperator).State = 2
End Select
UI_Set_BG _RGB32(0, 191, 0)
'----------------------------
'-------- Side Pane --------
UI_Side_Pane = UI_New("UI_Side_Pane", UI_TYPE_Frame, 0, 16, 128, -16, "", "")
UI_Set_BG _RGB32(63)
'---------------------------
'-------- Open File Dialog --------
UI_Dialog_OpenFile = UI_New("UI_Dialog_OpenFile", UI_TYPE_Dialog, 0, 0, -128, -128, "Open File", "")
UI_Set_FG -1
UI(UI_Dialog_OpenFile).Visible = 0
'----------------------------------
'-------- Save As Dialog --------
UI_Dialog_SaveFileAs = UI_New("UI_Dialog_SaveFileAs", UI_TYPE_Dialog, 0, 0, -128, -128, "Save As File", "")
UI_Set_FG -1
UI(UI_Dialog_SaveFileAs).Visible = 0
'--------------------------------

Dim Shared As Long MainScreen, tmpScreen
MainScreen = _NewImage(960, 540, 32)
Screen MainScreen
Do Until _ScreenExists: Loop
While _Resize: Wend

Color Config_TextColor, 0

If _CommandCount Then
        For I = 1 To _CommandCount
                If _FileExists(Command$(I)) Then INFILE$ = Command$(I) Else INFILE$ = _StartDir$ + FILE_SEPERATOR + Command$(I)
                OpenFile INFILE$
        Next I
Else
        OpenWorkspace Config_CurrentWorkspace
End If

Dim Shared As Long TextOffsetX, TextOffsetY: TextOffsetY = 16

Do
        If _Resize Then
                tmpW = _ResizeWidth: tmpH = _ResizeHeight
                If tmpW > 0 And tmpH > 0 Then tmpScreen = MainScreen: MainScreen = _NewImage(tmpW, tmpH, 32): Screen MainScreen: _FreeImage tmpScreen
        End If

        Cls , Config_BackgroundColor
        Color Config_TextColor, 0
        _Limit 60
        CurrentFile = Clamp(LBound(File), CurrentFile, UBound(File))
        If CurrentFile Then SetTitle File(CurrentFile).Name + " - Code Editor" Else SetTitle "Code Editor"
        While _MouseInput
                File(CurrentFile).ScrollOffset = Clamp(1, File(CurrentFile).ScrollOffset + 3 * _MouseWheel, File(CurrentFile).TotalLines)
        Wend

        OpenFileTasks: SaveFileTasks

        K$ = InKey$
        KeyHit = _KeyHit
        KeyShift = _KeyDown(100303) Or _KeyDown(100304)
        KeyCtrl = _KeyDown(100305) Or _KeyDown(100306)
        KeyAlt = _KeyDown(100307) Or _KeyDown(100308)
        DrawMenuBar
        DrawStatusBar
        If CurrentFile Then
                TextOffsetX = UI(UI_Side_Pane).W + 8 * ceil(Log(Max(1, File(CurrentFile).TotalLines)) / Log(10) + 2)
                DrawText
        End If
        UI_Draw
        If UI_Focus = 0 Then
                If _MouseButton(1) And MouseInBox(0, 16, _Width - 9, _Height - 17) Then
                        If KeyAlt Then AddCursor _SHR(_MouseX - TextOffsetX, 3) + 1, _SHR(_MouseY - TextOffsetY, 4) + 1 Else SetCursor _SHR(_MouseX - TextOffsetX, 3) + 1, _SHR(_MouseY - TextOffsetY, 4) + 1
                End If
                For I = 1 To Len(File(CurrentFile).Cursors) Step 8
                        CursorX = CVL(Mid$(File(CurrentFile).Cursors, I, 4))
                        CursorY = CVL(Mid$(File(CurrentFile).Cursors, I + 4, 4))
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        If RopeI = 0 Then _Continue
                        BoundCursor = False
                        If Len(K$) Then
                                Select Case Asc(K$)
                                        Case 8: DeleteText 1, CursorX, CursorY: BoundCursor = True
                                        Case 13: If CursorX = Len(Rope(RopeI)) + 1 Then
                                                        InsertLine CursorY + 1
                                                        CursorX = 1
                                                        CursorY = CursorY + 1
                                                Else
                                                        InsertLine CursorY + 1
                                                        T$ = Mid$(Rope(RopeI), CursorX)
                                                        Rope(RopeI) = Left$(Rope(RopeI), CursorX - 1)
                                                        CursorY = CursorY + 1: CursorX = 1
                                                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                                                        Rope(RopeI) = T$
                                                End If
                                        Case 32 To 126: InsertText K$, CursorX, CursorY: BoundCursor = True
                                End Select
                        End If
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        ShowCursor = Timer(0.1) - Int(Timer(0.1)) < 0.5
                        If BoundCursor Then
                                CursorX = Clamp(1, CursorX, Len(Rope(RopeI)) + 1)
                                CursorY = Clamp(1, CursorY, File(CurrentFile).TotalLines)
                                ShowCursorTime = Timer(0.1) + 1
                        End If
                        Select Case KeyHit
                                Case 19200 'Left
                                        CursorX = Max(CursorX - 1, 1)
                                Case 19712 'Right
                                        CursorX = Min(CursorX + 1, Len(Rope(RopeI)) + 1)
                                Case 18432 'Up
                                        CursorY = Max(CursorY - 1, 1)
                                Case 20480 'Down
                                        CursorY = Min(CursorY + 1, File(CurrentFile).TotalLines)
                                Case 18176 'Home
                                        CursorX = 1
                                Case 20224 'End
                                        CursorX = Len(Rope(RopeI)) + 1
                        End Select
                        Mid$(File(CurrentFile).Cursors, I, 4) = MKL$(CursorX)
                        Mid$(File(CurrentFile).Cursors, I + 4, 4) = MKL$(CursorY)
                        'Display Cursor
                        If ShowCursor = 0 And Timer(0.1) > ShowCursorTime Then _Continue
                        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
                        X = TextOffsetX + _SHL(Min(CursorX, Len(Rope(RopeI)) + 1) - 1, 3)
                        Y = TextOffsetY + _SHL(CursorY - File(CurrentFile).ScrollOffset, 4)
                        Line (X, Y)-(X + 7, Y + 15), -1, B
                Next I
        End If
        _Display
        If UI(UI_MenuButton_File).Response = 1 Or (KeyCtrl And (KeyHit = 78 Or KeyHit = 110)) Then NewFile
        If UI(UI_MenuButton_File).Response = 2 Or (KeyCtrl And (KeyHit = 79 Or KeyHit = 111)) Then UI(UI_Dialog_OpenFile).Visible = -1
        If UI(UI_MenuButton_File).Response = 3 Or (KeyCtrl And (KeyHit = 83 Or KeyHit = 115)) Then SaveFile CurrentFile
        If UI(UI_MenuButton_File).Response = 4 Then
                UI(UI_Dialog_SaveFileAs).Visible = -1
        End If
        If UI(UI_MenuButton_File).Response = 5 Or (KeyCtrl And (KeyHit = 87 Or KeyHit = 119)) Then
                EmptyRopePoints = EmptyRopePoints + File(CurrentFile).Content
                For I = CurrentFile + 1 To UBound(File): Swap File(I), File(I - 1): Next I
                If I > 2 Then ReDim _Preserve File(1 To I - 2) As File Else ReDim File(0) As File
        End If
        If UI(UI_MenuButton_File).Response = 6 Or (KeyCtrl And (KeyHit = 81 Or KeyHit = 113)) Then Exit Do
        If UI(UI_MenuButton_View).Response = 3 Or (KeyAlt And (KeyHit = 84 Or KeyHit = 116)) Then UI(UI_Side_Pane).W = 144 - UI(UI_Side_Pane).W
        Select Case UI(UI_ToggleButton_LineSeperator).Response
                Case 1: Config_FileSeperator = Chr$(13) + Chr$(10)
                Case 2: Config_FileSeperator = Chr$(10)
                Case 3: Config_FileSeperator = Chr$(13)
        End Select
        If _Exit Then Exit Do
Loop
WriteConfigFile
System

Sub AddCursor (X As Long, Y As Long)
        For I = 1 To Len(File(CurrentFile).Cursors) Step 8
                CX = CVL(Mid$(File(CurrentFile).Cursors, I, 4))
                CY = CVL(Mid$(File(CurrentFile).Cursors, I + 4, 4))
                If CX = X And CY = Y Then Exit Sub
        Next I
        File(CurrentFile).Cursors = File(CurrentFile).Cursors + MKL$(X) + MKL$(Y)
End Sub
Sub SetCursor (X As Long, Y As Long)
        File(CurrentFile).Cursors = MKL$(X) + MKL$(Y)
End Sub

Sub InsertText (T$, CursorX As Long, CursorY As Long)
        RopeI = CVL(Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4))
        Rope(RopeI) = Left$(Rope(RopeI), CursorX - 1) + T$ + Mid$(Rope(RopeI), CursorX)
        CursorX = CursorX + 1
End Sub
Sub InsertLine (CursorY As Long)
        File(CurrentFile).TotalLines = File(CurrentFile).TotalLines + 1
        File(CurrentFile).Content = Left$(File(CurrentFile).Content, _SHL(CursorY - 1, 2)) + MKL$(GetNewRopePointer) + Mid$(File(CurrentFile).Content, _SHL(CursorY - 1, 2) + 1)
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
                        CursorX = CursorX - 1
                End If
        Next I
End Sub
Sub DeleteLine (CursorY As Long)
        EmptyRopePoints = EmptyRopePoints + Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) - 3, 4)
        File(CurrentFile).Content = Left$(File(CurrentFile).Content, _SHL(CursorY - 1, 2)) + Mid$(File(CurrentFile).Content, _SHL(CursorY, 2) + 1)
        File(CurrentFile).TotalLines = File(CurrentFile).TotalLines - 1
End Sub

Sub DrawMenuBar
        Line (0, 0)-(_Width - 1, 16), _RGB32(0, 63, 127), BF
        If CurrentFile Then _PrintString (_Width - 8 * Len(File(CurrentFile).Name) - 16, 0), "[" + File(CurrentFile).Name + "]"
End Sub
Sub DrawStatusBar
        Line (0, _Height - 16)-(_Width - 1, _Height - 1), _RGB32(0, 63, 127), BF
End Sub
Sub DrawText
        TotalLinesVisible = _SHR(_Height - 32, 4) - 1
        HorizontalCharsVisible = _SHR(_Width - TextOffsetX, 3)
        J = TextOffsetY
        For I = File(CurrentFile).ScrollOffset To Min(File(CurrentFile).ScrollOffset + TotalLinesVisible, File(CurrentFile).TotalLines)
                K = CVL(Mid$(File(CurrentFile).Content, I * 4 - 3, 4))
                L$ = _Trim$(Str$(I))
                L$ = Space$(_SHR(TextOffsetX, 3) - Len(L$) - 1) + L$ + Chr$(179) + Rope(K)
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
        Config_CurrentWorkspace = ""
        Config_FileSeperator = Chr$(13) + Chr$(10)
End Sub
Sub CreateConfig
        MapSetKey Config$, "BackgroundColor", MKL$(Config_BackgroundColor)
        MapSetKey Config$, "TextColor", MKL$(Config_TextColor)
        MapSetKey Config$, "CurrentWorkspace", Config_CurrentWorkspace
        MapSetKey Config$, "FileSeperator", Config_FileSeperator
End Sub
Sub ParseConfig
        Config_BackgroundColor = CVL(MapGetKey(Config$, "BackgroundColor"))
        Config_TextColor = CVL(MapGetKey(Config$, "TextColor"))
        If Len(Config_CurrentWorkspace) Then OpenWorkspace Config_CurrentWorkspace
        Config_FileSeperator = MapGetKey(Config$, "FileSeperator")
End Sub
Sub WriteConfigFile
        CreateConfig
        If _FileExists("config.ini") Then Kill "config.ini"
        F = FreeFile
        Open "config.ini" For Binary As #F
        Put #F, , Config$
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
End Sub
Sub OpenFileTasks: Static As Long OpeningFileDialog, OpeningFile_FileNameLabel, OpeningFile_Progress
        If OpeningFileDialog = 0 Then
                OpeningFileDialog = UI_New("Dialog_OpeningFile", UI_TYPE_Dialog, 0, 0, 256, 64, "Opening File", "")
                UI(OpeningFileDialog).Visible = 0
                UI_PARENT = OpeningFileDialog
                OpeningFile_FileNameLabel = UI_New("Dialog_OpeningFile_Label", UI_TYPE_Label, 16, 24, 224, 16, "", "")
                OpeningFile_Progress = UI_New("Dialog_OpeningFile_Progress", UI_TYPE_ProgressBar, 16, 48, 224, 4, "", "")
                UI_PARENT = 0
        End If

        UI(OpeningFileDialog).Visible = 0
        For FileID = 1 To UBound(File)
                If File(FileID).Opened Then _Continue
                For J = 1 To 64
                        UI(OpeningFile_FileNameLabel).Label = File(FileID).Name
                        UI(OpeningFileDialog).Visible = -1
                        Line Input #File(FileID).FileID, L$
                        I = GetNewRopePointer
                        Rope(I) = L$
                        File(FileID).TotalLines = File(FileID).TotalLines + 1
                        If Len(File(FileID).Content) < File(FileID).TotalLines * 4 + 4 Then File(FileID).Content = File(FileID).Content + String$(1024, 0)
                        Mid$(File(FileID).Content, File(FileID).TotalLines * 4 - 3, 4) = MKL$(I)
                        UI(OpeningFile_Progress).Progress = Seek(File(FileID).FileID) / LOF(File(FileID).FileID)
                        If Seek(File(FileID).FileID) > LOF(File(FileID).FileID) Then Close #File(FileID).FileID: File(FileID).Opened = 1: Exit For
                Next J
        Next FileID
End Sub
Sub SaveFile (I As Long)
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
                For J = 1 To 64
                        File(FileID).SaveOffset = File(FileID).SaveOffset + 1
                        If File(FileID).SaveOffset > File(FileID).TotalLines Then Exit For
                        RopeI = CVL(Mid$(File(FileID).Content, _SHL(File(FileID).SaveOffset, 2) - 3, 4))
                        If File(FileID).SaveOffset = File(FileID).TotalLines Then Print #File(FileID).FileID, Rope(RopeI); Else Print #File(FileID).FileID, Rope(RopeI) + Config_FileSeperator;
                Next J
                UI(SavingFile_Progress).Progress = File(FileID).SaveOffset / File(FileID).TotalLines
        End If
        If File(FileID).SaveOffset < File(FileID).TotalLines Then Exit Sub
        SaveFileQueue = Mid$(SaveFileQueue, 5)
        Close #File(FileID).FileID
End Sub
Function GetNewRopePointer&
        If Len(EmptyRopePoints) = 0 Then
                I = UBound(Rope) + 1
                ReDim _Preserve Rope(1 To I) As String
                GetNewRopePointer& = I
        Else
                GetNewRopePointer& = CVL(Left$(EmptyRopePoints, 4))
                Rope(CVL(Left$(EmptyRopePoints, 4))) = ""
                EmptyRopePoints = Mid$(EmptyRopePoints, 5)
        End If
End Function
'---------------------------------

'-------- Workspace Management --------
Sub OpenWorkspace (Path$)
End Sub
Sub SaveWorkspace
End Sub
Sub CloseWorkspace
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
        T~& = Max(_InStrRev(Path$, "/"), _InStrRev(Path$, "\"))
        If T~& = 0 Then GetLastPathName$ = Path$ Else GetLastPathName$ = Mid$(Path$, T~&)
End Function
Function GetKeyDown (__Key As _Unsigned Long)
        If __Key = 0 Then Exit Function
        GetKeyDown = _KeyDown(__Key)
End Function
Function ceil& (t#)
        t& = Int(t#)
        ceil& = t& + Sgn(t# - t&)
End Function

Function GetDirList$ (CurrentDirectory$)
        If CurrentDirectory$ = "" Then
                $If WIN Then
                        O$ = ListStringNew
                        For I = 65 To 90
                                If _DirExists(Chr$(I) + ":\") Then ListStringAdd O$, Chr$(I) + ":"
                        Next I
                        GetDirList$ = O$
                        Exit Function
                $Else
                                CurrentDirectory$ = "/"
                $End If
        End If
        O$ = ListStringNew
        Shell "dir /b /o:n " + Chr$(34) + CurrentDirectory$ + Chr$(34) + " > tmp.txt"
        __F = FreeFile
        Open "tmp.txt" For Input As __F
        Do
                Line Input #__F, L$
                ListStringAdd O$, L$
        Loop Until EOF(__F)
        Close #__F
        Kill "tmp.txt"
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
                Select Case __UI.Type And __UI.Visible
                        Case UI_TYPE_Frame
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, B

                        Case UI_TYPE_Label
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, B
                                PrintString __UI.__X, __UI.__Y + (__UI.__H - _FontHeight) / 2, __UI.Label

                        Case UI_TYPE_Button
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), _RGB32(255, 127), B
                                If UI_Focus = __UI.Parent And MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) Then
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
                                        Line (__UI.__X, __UI.__Y + __UI.__H)-(__UI.__X + __UI.MenuDialogWidth, __UI.__Y + __UI.__H + __UI.MenuDialogHeight), _RGB32(0, 191), BF
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
                                If UI_Focus = __UI.Parent And MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = I: WaitForMouseButtonRelease
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.BG, BF
                                __H = Clamp(16, __UI.__H / __UI.MaxScroll, __UI.__H)
                                __Y = (__UI.__H - __H) * (__UI.Scroll - 1) / __UI.MaxScroll
                                Line (__UI.__X + 1, __UI.__Y + __Y)-(__UI.__X + __UI.__W - 1, __UI.__Y + __Y + __H), _RGB32(255), BF
                                If UI_Focus = __UI.Parent And MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then __UI.Scroll = (_MouseY - __UI.__Y - 8) * __UI.MaxScroll / __UI.__H

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
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = __UI.Parent: WaitForMouseButtonRelease
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), __UI.FG, B
                                __UI.Scroll = Clamp(1, __UI.Scroll + UI_MouseWheel, ListStringLength(__UI.ParsedContent))
                                K = 0
                                For J = Max(1, __UI.Scroll) To Min(__UI.Scroll + __UI.__H / _FontHeight - 1, ListStringLength(__UI.ParsedContent))
                                        X1 = __UI.__X + _FontWidth
                                        Y1 = __UI.__Y + K
                                        X2 = __UI.__X + __UI.__W - _FontWidth
                                        Y2 = __UI.__Y + K + _FontHeight - 1
                                        If MouseInBox(X1, Y1, X2, Y2) Then
                                                Line (X1, Y1)-(X2, Y2), __UI.BG, BF
                                                If _MouseButton(1) Then
                                                        WaitForMouseButtonRelease
                                                        __UI.Selected = J
                                                        __UI.Response = J
                                                End If
                                        End If
                                        If __UI.Selected = J Then Line (X1, Y1)-(X2, Y2), __UI.BG, BF
                                        PrintString X1, Y1, ListStringGet(__UI.ParsedContent, J)
                                        K = K + _FontHeight
                                Next J

                        Case UI_TYPE_TextView
                                Line (__UI.__X, __UI.__Y)-(__UI.__X + __UI.__W, __UI.__Y + __UI.__H), _RGB32(255), B
                                If MouseInBox(__UI.__X, __UI.__Y, __UI.__X + __UI.__W, __UI.__Y + __UI.__H) And _MouseButton(1) Then UI_Focus = I: WaitForMouseButtonRelease
                                If UI_Focus = I Then
                                        Select Case KeyHit
                                                Case 8: __UI.Content = Left$(__UI.Content, Len(__UI.Content) - 1)
                                                Case 9: UI_Focus = __UI.Parent
                                                Case 32 To 122: __UI.Content = __UI.Content + Chr$(KeyHit)
                                        End Select
                                        __UI.Response = KeyHit
                                End If
                                PrintString __UI.__X, __UI.__Y, __UI.Content + "_"

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
