Attribute VB_Name = "ModPas"
Enum WhichColour
    BackColour = 0
    FColour = 1
End Enum

Public CurrentXpos As Integer, CurrentYPos As Integer
Public CLColour As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Function RestoreOld(window As PictureBox)
    CurrentXpos = 0
    CurrentYPos = 0
    window.FontSize = 9.75
    window.ForeColor = vbWhite
End Function
Function Delay(NumSec As Integer)
Dim Max As Integer
    Max = 100
        If NumSec = 0 Then
            Exit Function
        Else
            Sleep NumSec * Max
        End If
        
End Function
Function SetTextPositionsX(lzString As String, Ln As Integer)
Dim StrVal As String
    StrVal = Trim(Left(lzString, Len(lzString) - 1))
    If Len(StrVal) = 0 Then
        GetLastError 8, Ln
        Exit Function
    Else
        CurrentXpos = Val(StrVal)
    End If
    StrVal = ""
    
End Function

Function SetTextPositionsY(lzString As String, Ln As Integer)
Dim StrVal As String
    StrVal = Trim(Left(lzString, Len(lzString) - 1))
    If Len(StrVal) = 0 Then
        GetLastError 8, Ln
        Exit Function
    Else
        CurrentYPos = Val(StrVal)
    End If
    StrVal = ""
    
End Function

Function FindPart(lzStr As String, mPart As String) As Integer
Dim TPos As Integer
    TPos = InStr(lzStr, mPart)
    If TPos Then
        FindPart = 1
    Else
        FindPart = 0
    End If
    
End Function
Function FindPoint(lzStr As String, mPart As String) As Integer
Dim Xpos As Integer
    Xpos = InStr(lzStr, mPart)
    If Xpos > 0 Then
        FindPoint = Xpos
    Else
        FindPoint = 0
    End If
    
End Function

Function ShowMsg(MsgText As String)
Dim MsgStyle As String
    MsgStyle = Trim(Right(MsgText, 1))
    
    If IsDigit(MsgStyle) = False Then
        MsgBox MsgText
    Else
        If Val(MsgStyle) > 4 Or Val(MsgStyle) < 1 Then
            GetLastError 13, 0
            Exit Function
        Else
        
        Select Case MsgStyle
            Case 1
                MsgBox Left(MsgText, Len(MsgText) - 2), vbCritical
            Case 2
                MsgBox Left(MsgText, Len(MsgText) - 2), vbExclamation
            Case 3
                MsgBox Left(MsgText, Len(MsgText) - 2), vbInformation
            Case 4
                MsgBox Left(MsgText, Len(MsgText) - 2), vbQuestion
            End Select
        End If
    End If
    
End Function
Function IsDigit(ByVal Digit As String) As Boolean
Dim Counter As Integer
    For Counter = 1 To Len(Digit)
        ch = Mid(Digit, Counter, 1)
        If ch Like "[0-9]" Then
            IsDigit = True
        Else
            IsDigit = False
        End If
    Next
    Counter = 0
    
End Function
Function ScreenModes(TMode As Integer, window As PictureBox)
    Select Case TMode
        Case 10
            window.FontSize = 5
        Case 12
            window.FontSize = 16
        Case 13
            window.FontSize = 18
        Case 16
            window.FontSize = 20
            
        Case Else
        MsgBox "Mode " & TMode & " Is Not Sopprted in this verision", vbInformation
        
    End Select
    
End Function
Function GetLastError(ErrorNum As Integer, LineNum As Integer)
    Select Case ErrorNum
        Case 1
            MsgBox "Program with out PROGRAM at Line " & LineNum
        Case 2
            MsgBox "Expected ; not found At Line " & LineNum
        Case 3
            MsgBox "Expected Mode with out value At Line " & LineNum
        Case 4
            MsgBox "Expected Text Colour without = At Line " & LineNum
        Case 5
            MsgBox "Expected ( missing in Statement at Line " & LineNum
        Case 6
            MsgBox "Expected ) missing in statement at line " & LineNum
        Case 7
            MsgBox "Expected , missing in statement at line " & LineNum
        Case 8
            MsgBox "Expected Value missing in function at line " & LineNum
        Case 9
            MsgBox "Invaild data value entered at line " & LineNum
        Case 10
            MsgBox "Expected END. not found at line " & LineNum
        Case 11
            MsgBox "Expected = not found at line " & LineNum
        Case 12
            MsgBox "Invaild Ellipse Data was entered at line " & LineNum
        Case 13
            MsgBox "Invaild Mesaage Box Style const 1 to 3 are only allowed in this verision", vbInformation
            
        End Select
        
End Function
Function SetColour(TColour As String, ColourType As WhichColour, window As PictureBox)
    Select Case UCase(TColour)
        Case "CLRED"
            CLColour = vbRed
        Case "CLBLUE"
            CLColour = vbBlue
        Case "CLGREEN"
            CLColour = vbGreen
        Case "CLBLACK"
            CLColour = vbBlack
        Case "CLYELLOW"
            CLColour = vbYellow
        Case "CLWHITE"
            CLColour = vbWhite
        Case "CLDESKTOP"
            CLColour = vbDesktop
        Case "CLCYAN"
            CLColour = vbCyan
        Case "CLMAGENTA"
            CLColour = vbMagenta
        Case Else
            MsgBox TColour & " Is not sopported in this verision", vbInformation
        End Select
        
        If ColourType = BackColour Then
            window.BackColor = CLColour
        ElseIf ColourType = FColour Then
            window.ForeColor = CLColour
        End If
        
        
End Function
Function GetText(mText As String) As String
Dim Lpos As Integer
Dim StrL As String

    Lpos = InStr(mText, "(")
    StrL = Mid(mText, Lpos + 1, InStr(Lpos + 1, mText, ")") - Lpos - 1)
    StrL = Replace(StrL, Chr(34), "")
    GetText = StrL
    
    
End Function
Function PutToScreen(lzStr As String, window As PictureBox)
    If CurrentXpos = 0 Or CurrentYPos = 0 Then
        window.Print lzStr & vbclrf
        Exit Function
    Else
        window.CurrentX = CurrentXpos
        window.CurrentY = CurrentYPos
        window.Print lzStr & vbclrf
    End If
    
End Function
Function RemoveChar(StrString As String, SChar As String) As String
    RemoveChar = Replace(StrString, Chr(9), Chr(32))
    
End Function

Function Plot(lzStr As String, Ln As Integer, window As PictureBox)
Dim StrVal1, StrVal2 As String
Dim StrVal As String
Dim k As String
Dim Val1, Val2 As Integer

    k = lzStr
    
    If FindPart(k, "(") = 0 Then
        GetLastError 5, Ln
        Exit Function
    ElseIf FindPart(k, ")") = 0 Then
        GetLastError 6, Ln
        Exit Function
    ElseIf FindPart(k, ";") = 0 Then
        GetLastError 2, Ln
        Exit Function
    Else
        StrVal = GetText(k)
        If FindPart(k, ",") = 0 Then
            GetLastError 7, Ln
            Exit Function
        Else
            StrVal = Trim(GetText(k))
            StrVal1 = Trim(Mid(StrVal, FindPoint(StrVal, ",") + 1, Len(StrVal)))
            StrVal2 = Trim(Mid(StrVal, 1, FindPoint(StrVal, ",") - 1))
            
            If IsDigit(StrVal1) = False Then
                GetLastError 9, Ln
                Exit Function
            ElseIf IsDigit(StrVal2) = False Then
                Exit Function
                GetLastError 9, Ln
            Else
                Val1 = Val(StrVal2)
                Val2 = Val(StrVal1)
                window.PSet (Val1, Val2), CLColour
            End If
    End If
    End If
    k = ""
    StrVal1 = ""
    StrVal2 = ""
    Val1 = 0
    Val2 = 0
    
End Function
Function DrawEllipse(lzStr As String, Ln As Integer, window As PictureBox)
Dim EllipseData As Collection
Dim Lpos As Integer
Dim StrVal As String, G As String

    Set EllipseData = New Collection
    StrVal = GetText(lzStr)
    If Len(StrVal) = 0 Then
        GetLastError 8, Ln
        Exit Function
    Else
        StrVal = Left(StrVal, Len(StrVal)) & ","
        For Lpos = 1 To Len(StrVal)
            ch = Mid(StrVal, Lpos, 1)
            G = G & ch
            If InStr(G, ",") Then
                mCount = mCount + 1
                
                G = Left(G, Len(G) - 1)
                If IsDigit(G) = False Then
                    GetLastError 9, Ln
                    Exit Function
                Else
                   EllipseData.Add G
                G = ""
                End If
            End If
        Next
    End If
    
    If EllipseData.Count = 3 Then
        window.Circle (EllipseData(1), EllipseData(2)), EllipseData(3), CLColour
        
    Else
        GetLastError 12, Ln
    End If
        StrVal = ""
        G = ""
        mCount = 0
        
End Function
Function DrawLine(lzStr As String, Ln As Integer, window As PictureBox)
Dim StrVal1, StrVal2 As String
Dim StrVal As String
Dim k As String
Dim Val1, Val2 As Integer

    k = lzStr
    
    If FindPart(k, "(") = 0 Then
        GetLastError 5, Ln
        Exit Function
    ElseIf FindPart(k, ")") = 0 Then
        GetLastError 6, Ln
        Exit Function
    ElseIf FindPart(k, ";") = 0 Then
        GetLastError 2, Ln
        Exit Function
    Else
        StrVal = GetText(k)
        If FindPart(k, ",") = 0 Then
            GetLastError 7, Ln
            Exit Function
        Else
            StrVal = Trim(GetText(k))
            StrVal1 = Trim(Mid(StrVal, FindPoint(StrVal, ",") + 1, Len(StrVal)))
            StrVal2 = Trim(Mid(StrVal, 1, FindPoint(StrVal, ",") - 1))
            
            If IsDigit(StrVal1) = False Or IsDigit(StrVal2) = False Then
                GetLastError 9, Ln
                Exit Function
            Else
                Val1 = Val(StrVal2)
                Val2 = Val(StrVal1)
                LineTo window.hdc, Val1, Val2
            End If
    End If
    End If
    k = ""
    StrVal1 = ""
    StrVal2 = ""
    Val1 = 0
    Val2 = 0
    
    
End Function
