VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM++ Compiler Beta 2"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6420
      Top             =   3855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1108
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":154A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2210
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2652
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3318
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   688
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find Text"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   15
      MouseIcon       =   "Form1.frx":375A
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":3B9C
      Top             =   495
      Width           =   7470
   End
   Begin VB.PictureBox PCode 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   90
      ScaleHeight     =   3855
      ScaleWidth      =   7305
      TabIndex        =   1
      Top             =   5010
      Visible         =   0   'False
      Width           =   7365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      Index           =   1
      X1              =   -315
      X2              =   1515
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   -315
      X2              =   1515
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Project"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Project"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindTxt 
         Caption         =   "&Find Text"
      End
   End
   Begin VB.Menu mnuComp1 
      Caption         =   "&Compile"
      Begin VB.Menu mnuComp 
         Caption         =   "C&ompile"
      End
      Begin VB.Menu mnuStopComp 
         Caption         =   "&Stop Compile"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGuid 
         Caption         =   "&See Help Guid"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About DM++"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyWords(1 To 18) As String
Dim TCount As Integer
Dim CarryOn As Boolean
Dim RemmberStuff As String
Sub SaveProject()
Dim Data As String, FExt, Filename As String
Dim Ans

    Filename = RemoveNulls(SaveFile)
    FExt = Right(UCase(Filename), 2)
   
    If Len(Filename) = 0 Then
        Exit Sub
    Else
        If FExt = "DM" Then
    Else
        Filename = Filename & ".dm"
    End If
        Data = Text1.Text
        If Not Dir(Filename) = "" Then
            Ans = _
                MsgBox("Do you want to replace this file", _
                vbYesNo)
                If Ans = vbNo Then
                    Exit Sub
                Else
                    Kill Filename
            End If
        End If
        
        Open Filename For Binary As #1
        Put #1, , Data
        Close #1
    End If
    Filename = ""
    FExt = ""
    Data = ""
    
End Sub
Sub OpenProject()
Dim Data As String, Filename, FExt As String
Dim Filenum As Long

    Filenum = FreeFile
    Filename = RemoveNulls(OpenFile)
    If Len(Filename) = 0 Then
        Exit Sub
    Else
        FExt = UCase(Right(Filename, 2))
        If FExt = "DM" Then
            Open Filename For Binary As #Filenum
            Data = Space(LOF(Filenum))
            Get #Filenum, , Data
            Close #Filenum
            Text1.Text = Data
        Else
            MsgBox "This is not a viald DM++ project.", vbCritical, "Error...."
            Exit Sub
        End If
    End If
    Filename = ""
    FExt = ""
    Data = ""
        
End Sub
Function RemoveChar(StrString As String, SChar As String) As String
Dim Xpos As Integer
    Xpos = InStr(StrString, SChar)
    If Xpos Then
        RemoveChar = Left(StrString, Xpos - 1)
    Else
        RemoveChar = StrString
    End If
    Xpos = 0
    
End Function

Function TCompile(lzStr As String)
Dim LineNum  As Integer, I As Integer
Dim StrBuff As String, TCode As String
    
    For I = 1 To Len(lzStr)
        ch = Asc(Mid(lzStr, I, 1))
        If ch <> 13 Then
            StrBuff = StrBuff & Mid(lzStr, I, 1)
        Else
            LineNum = LineNum + 1
            StrBuff = Trim(StrBuff)
            'RemoveChar StrBuff, Chr(9)
            If InStr(StrBuff, KeyWords(1)) And LineNum = 1 Then
                CarryOn = True
            End If
            
            If CarryOn = False Then
                GetLastError 1, LineNum
                Exit Function
            Else
            End If
            '//-------------------------------------------------------------------
            
            If InStr(StrBuff, KeyWords(2)) Then
                TCount = TCount + 1
                If TCount > 0 Then
                    If FindPart(StrBuff, ";") = 0 Then
                        GetLastError 2, LineNum
                        Exit Function
                    Else
                    TCode = StrBuff
                    PutToScreen GetText(TCode), PCode
                    End If
                End If
            End If
            TCount = 0
            TCode = ""
            '//-------------------------------------------------------------------
            If InStr(StrBuff, KeyWords(3)) Then
                TCount = TCount + 1
                If TCount > 0 Then
                    If FindPart(StrBuff, ";") = 0 Then
                        GetLastError 1, LineNum
                        Exit Function
                    Else
                    TCode = StrBuff
                    TCode = GetText(TCode)
                    If Len(TCode) = 0 Then
                        GetLastError 3, LineNum
                        Exit Function
                    Else
                        If IsDigit(TCode) = False Then
                            GetLastError 9, LineNum
                            Exit Function
                        Else
                        ScreenModes Val(TCode), PCode
                    End If
                    End If
                End If
            End If
        End If
                '//-------------------------------------------------------------------

        TCount = 0
        TCode = ""
      
      If InStr(StrBuff, KeyWords(4)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                TCol = FindPoint(TCode, "=")
                If TCol > 0 Then
                    TCode = Mid(TCode, TCol + 1, Len(TCode))
                    TCode = Left(TCode, Len(TCode) - 1)
                    SetColour TCode, FColour, PCode
                Else
                    GetLastError 4, LineNum
                    Exit Function
                End If
            End If
        End If
    End If
            '//-------------------------------------------------------------------

    TCode = ""
    TCount = 0
    If InStr(StrBuff, KeyWords(5)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                If FindPoint(TCode, "(") = 0 Then
                    GetLastError 5, LineNum
                ElseIf FindPoint(TCode, ")") = 0 Then
                    GetLastError 6, LineNum
                    Exit Function
                    Else
                        TCode = GetText(TCode)
                        ShowMsg TCode
                End If
            End If
        End If
    End If
    TCount = 0
    TCode = ""
            '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(6)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            Plot StrBuff, LineNum, PCode
        End If
    End If
            '//-------------------------------------------------------------------
            
    If InStr(StrBuff, KeyWords(7)) Then
        TCount = TCount + 1
        If TCount > 0 Then
           If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                Beep
            End If
        End If
    End If
    TCount = 0
            '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(8)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 2, LineNum
                Exit Function
            Else
                PCode.Cls
            End If
        End If
    End If
    TCount = 0
                '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(9)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            DrawLine StrBuff, LineNum, PCode
        End If
    End If
    TCount = 0
                '//-------------------------------------------------------------------
                
      If InStr(StrBuff, KeyWords(11)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                TCol = FindPoint(TCode, "=")
                If TCol > 0 Then
                    TCode = Mid(TCode, TCol + 1, Len(TCode))
                    TCode = Left(TCode, Len(TCode) - 1)
                    SetColour TCode, BackColour, PCode
                Else
                    GetLastError 4, LineNum
                    Exit Function
                End If
            End If
        End If
    End If
    TCode = ""
    TCount = 0
            '//-------------------------------------------------------------------

    If InStr(StrBuff, KeyWords(12)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
               TCode = StrBuff
                If FindPoint(TCode, "=") = 0 Then
                    GetLastError 11, LineNum
                    Exit Function
                Else
                    SetTextPositionsX Mid(TCode, FindPoint(TCode, "=") + 1, Len(TCode)), LineNum
                End If
            End If
        End If
    End If
    
    '//-------------------------------------------------------------------
    
        If InStr(StrBuff, KeyWords(13)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
               TCode = StrBuff
                If FindPoint(TCode, "=") = 0 Then
                    GetLastError 11, LineNum
                    Exit Function
                Else
                    SetTextPositionsY Mid(TCode, FindPoint(TCode, "=") + 1, Len(TCode)), LineNum
                End If
            End If
        End If
    End If
    TCode = ""
    TCount = 0
    
    If InStr(StrBuff, KeyWords(14)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
                TCode = StrBuff
                DrawEllipse TCode, LineNum, PCode
            End If
        End If
    End If
    TCode = ""
    TCount = 0
    
    '//-------------------------------------------------------------------
    
    If InStr(StrBuff, KeyWords(15)) Then
        TCount = TCount + 1
        If TCount > 0 Then
            If FindPart(StrBuff, ";") = 0 Then
                GetLastError 1, LineNum
                Exit Function
            Else
            TCode = StrBuff
            If FindPoint(TCode, "=") = 0 Then
                GetLastError 11, LineNum
                Exit Function
            Else
                TCode = Mid(TCode, FindPoint(TCode, "="), Len(TCode) - 2)
                TCode = Right(TCode, Len(TCode) - 1)
                TCode = Left(TCode, Len(TCode) - 1)
                If IsDigit(TCode) = False Then
                    GetLastError 9, LineNum
                    Exit Function
                Else
                    Delay Val(TCode)
            End If
            End If
        End If
    End If
    End If
    
    '//-------------------------------------------------------------------
        StrBuff = ""
        I = I + 1
    End If
    Next
    
    I = 0
    TCode = ""
    StrBuff = ""
    TCount = 0
    LineNum = 0
    
End Function



Private Sub Form_Load()
    KeyWords(1) = "#BEGIN"
    KeyWords(2) = "TextOut"
    KeyWords(3) = "Mode"
    KeyWords(4) = "TextColour"
    KeyWords(5) = "ShowMessage"
    KeyWords(6) = "Plot"
    KeyWords(7) = "Beep"
    KeyWords(8) = "Cls"
    KeyWords(9) = "DrawLine"
    KeyWords(10) = "END."
    
    ' New Functions Added as of 16/01/01
    
    KeyWords(11) = "BkColour"
    KeyWords(12) = "CurrentX"
    KeyWords(13) = "CurrentY"
    KeyWords(14) = "DrawEllipse"
    KeyWords(15) = "Delay"
    
    TCount = 1
    ModPas.CLColour = vbBlack
    Toolbar1.Buttons(10).Enabled = False
    
    PCode.Top = Text1.Top
    PCode.Left = Text1.Left
    PCode.Width = Text1.Width
    
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = ScaleWidth - 1
    Line1(1).X2 = ScaleWidth - 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form1: End
    
End Sub

Private Sub mnuAbout_Click()
    FrmAbout1.Show
    Form1.Hide
            
End Sub

Private Sub mnuAll_Click()
    EditMenu Text1, "SELALL"
    
End Sub

Private Sub mnuComp_Click()
    Text1.Visible = False
    PCode.Visible = True
    RemmberStuff = Text1
    Text1 = TCompile(Text1)
    mnuOpen.Enabled = False
    mnuSave.Enabled = False
    Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(10).Enabled = True
    Text1 = ""
    
End Sub

Private Sub mnuCopy_Click()
    EditMenu Text1, "COPY"
    
End Sub

Private Sub mnuCut_Click()
    EditMenu Text1, "CUT"
    
End Sub

Private Sub mnuEx_Click()
    Ans = _
    MsgBox("Do you want to exit this program now", _
    vbYesNo)
    If Ans = vbNo Then
        Exit Sub
    Else
        RemmberStuff = ""
        Unload Form1: End
    End If
            
End Sub

Private Sub mnuFindTxt_Click()
    EditMenu Text1, "FIND"
    
End Sub

Private Sub mnuGuid_Click()
Dim Path As String
    Path = App.Path
    If Right(Path, 1) = "\" Then
        Path = Path
    Else
        Path = Path & "\"
    End If
    Main.RunProgran hwnd, Path & "Help.txt", vsMaxSized
    Path = ""
    
End Sub

Private Sub mnuOpen_Click()
    OpenProject
    
End Sub

Private Sub mnuPaste_Click()
    EditMenu Text1, "PASTE"
    
End Sub

Private Sub mnuSave_Click()
    SaveProject
    
End Sub

Private Sub mnuStopComp_Click()
    Text1.Visible = True
    PCode.Visible = False
    PCode.Cls
    Text1.Text = RemmberStuff
    mnuOpen.Enabled = True
    mnuSave.Enabled = True
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = False
    RemmberStuff = ""
    RestoreOld PCode
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnuOpen_Click
        Case 2
           mnuSave_Click
        Case 4
            mnuCut_Click
        Case 5
            mnuCopy_Click
        Case 6
            mnuPaste_Click
        Case 7
            mnuFindTxt_Click
        Case 9
            mnuComp_Click
        Case 10
            mnuStopComp_Click
        Case 11
           mnuAbout_Click
        Case 12
            mnuEx_Click
            
    End Select
        
End Sub

