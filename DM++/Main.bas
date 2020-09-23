Attribute VB_Name = "Main"
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Enum WinShow
    vsHide = 0
    vsNormal = 1
    vsMinSized = 2
    vsMaxSized = 3
End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function OpenFile() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.DM++ Files)" + Chr$(0) + "*.DM"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Open Project"
        ofn.FLAGS = 0
        
        A = GetOpenFileName(ofn)
        If (A) Then
                OpenFile = Trim(ofn.lpstrFile)
        End If
        
 End Function
 Public Function SaveFile() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.DM++ Files)" + Chr$(0) + "*.DM"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Save Project"
        ofn.FLAGS = 0
        
        A = GetSaveFileName(ofn)
        If (A) Then
                SaveFile = Trim(ofn.lpstrFile)
        End If
 End Function
Function RemoveNulls(lzString As String) As String
Dim Xpos As Integer
    Xpos = InStr(lzString, vbNullChar)
    If Xpos > 0 Then
        lzString = Left(lzString, Len(lzString) - 1)
        RemoveNulls = lzString
    End If
    
End Function
Function EditMenu(txtBox As TextBox, Cmd As String)
Dim StrFind As String
Dim Xpos As Integer

    Select Case Cmd
        Case "CUT"
            Clipboard.SetText txtBox.SelText
            txtBox.SelText = ""
        Case "COPY"
            Clipboard.SetText txtBox.SelText
        Case "PASTE"
            txtBox.SelText = Clipboard.GetText
        Case "SELALL"
            txtBox.SelStart = 0
            txtBox.SelLength = Len(txtBox.Text)
            
        Case "FIND"
            StrFind = InputBox("What do you want to find", "Find Text..", , 5, 5)
            If Len(StrFind) = 0 Then
                Exit Function
            Else
                Xpos = InStr(txtBox.Text, StrFind)
                If Xpos > 0 Then
                    txtBox.SetFocus
                    txtBox.SelStart = Xpos - 1
                    txtBox.SelLength = Len(StrFind)
                Else
                    Beep
                    MsgBox "Serach text " & Chr(34) & StrFind & Chr(34) & " was not found", vbExclamation
                End If
            End If
            Xpos = 0
            Cmd = ""
            StrFind = ""
    End Select
    
End Function
Public Function FileExists(ByVal Filename As String) As Integer
    If Dir(Filename) = "" Then FileExists = 0 Else FileExists = 1

End Function
Public Function RunProgran(mHwnd As Long, ProgramNamePath As String, ShowWindow As WinShow)
    If FileExists(ProgramNamePath) = 0 Then
        MsgBox "Can't find file " & ProgramNamePath, vbInformation
    Else
        ShellExecute mHwnd, vbNullString, ProgramNamePath, vbNullString, vbNullString, ShowWindow
    End If
    
End Function
