VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SparkLexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileName As String
Public Lines As List

Private CurPos As Long
Private CurChar As String
Private Chars As String
Private CurLine As Long
Private CurCol As Long
Private pCommentLevel As Integer

Private PlainCode As String

Private Const DIGITS = "0123456789"
Private Const LETTERS = "abcdefghijklmnopqrstuvwxyz"
Private Const SYMBOLS = "()[]{}!=+-*/%<>.,;&|^?:"
Private Const KEYWORDS = "_void__any__short__int__long__float__double__bool__string__const_" & _
                         "_true__false__return__break__continue__while__do__if__else__for_" & _
                         "_switch__case__default__private__public__import__typeof__throw_"

Public Event OnError(ErrMsg As String)

Public Function GetTokens(Code As String) As List
    Chars = Replace(Code, vbTab, "  ")
    Chars = Replace(Chars, vbNewLine, vbCr)
    PlainCode = Chars
    CurPos = 0
    
    Dim StrLines() As String
    Dim i As Long
    
    Set Lines = New List
    StrLines = Split(Chars, vbCr)
    
    For i = 0 To UBound(StrLines)
        Lines.Add StrLines(i)
    Next
    
    CurLine = 1
    NextChar
    
    Dim Tokens As New List
    
    Do While HasNextChar
        If pCommentLevel = 0 Then
            If CurChar = " " Or CurChar = vbCr Then
                If CurChar = vbCr Then
                    CurLine = CurLine + 1
                    CurCol = 0
                End If
                NextChar
            ElseIf CurChar = """" Or CurChar = "'" Then ' string
                Tokens.Add ParseString
            ElseIf InString(DIGITS, CurChar) Then ' number
                Tokens.Add ParseNumber
            ElseIf InString(LETTERS, LCase(CurChar)) Or CurChar = "_" Then ' identifier
                Tokens.Add ParseIdentifier
            ElseIf InString(SYMBOLS, CurChar) Then ' symbol
                Dim Token As Dictionary
                Set Token = ParseSymbol
                If Token Is Nothing = False Then Tokens.Add Token
            Else
                ThrowError "Unexpected token '" & CurChar & "'.", CurLine, CurCol
            End If
        Else
            If pCommentLevel = 1 And CurChar = vbCr Then
                pCommentLevel = 0
                CurLine = CurLine + 1
                CurCol = 0
            ElseIf pCommentLevel = 2 And CurChar = "*" Then
                NextChar
                If CurChar = "/" Then
                    pCommentLevel = 0
                Else
                    PrevChar
                End If
            End If
            NextChar
        End If
    Loop
    
    Set GetTokens = Tokens
End Function

Private Function IsCommentSymbol() As Boolean
    If CurPos < Len(Chars) - 1 Then
        If Mid(Chars, CurPos, 2) = "//" Then
            IsCommentSymbol = True
        End If
    End If
End Function

Private Function HasNextChar() As Boolean
    HasNextChar = CurPos <= Len(Chars)
End Function

Private Sub NextChar()
    CurPos = CurPos + 1
    CurCol = CurCol + 1
    CurChar = Mid(Chars, CurPos, 1)
End Sub

Private Sub PrevChar()
    CurPos = CurPos - 1
    CurCol = CurCol - 1
    CurChar = Mid(Chars, CurPos, 1)
End Sub

Private Function NewToken(TokenType As String, TokenValue As Variant) As Dictionary
    Set NewToken = New Dictionary
    NewToken.Add "type", TokenType
    NewToken.Add "value", TokenValue
    NewToken.Add "line", CurLine
    NewToken.Add "col", CurCol
End Function

Private Function ParseString() As Dictionary
    Dim Result As New Dictionary
    Dim Start As Long
    Start = CurPos
    Result.Add "type", "string"
    Result.Add "line", CurLine
    Result.Add "start", CurCol
    Dim Escape As Boolean
    Dim s As String
    Dim Quote As String
    Quote = CurChar
    NextChar
    Do
        If Escape Then
            Select Case CurChar
'            Case "\": s = s & "\"
'            Case """": s = s & """"
'            Case "'": s = s & "'"
            Case "n": s = s & vbNewLine 'vbLf
            Case "r": s = s & vbCr
            Case "t": s = s & vbTab
            Case Else
                s = s & CurChar
'                ThrowError "Invalid unescape character '" & CurChar & "'."
            End Select
            Escape = False
        ElseIf CurChar = "\" Then
            Escape = True
        ElseIf CurChar = Quote Then
            NextChar
            Exit Do
        Else
            s = s & CurChar
        End If
        If HasNextChar = False Then ThrowError "Unterminated string.", CurLine, Result("start")
        NextChar
    Loop
    Result.Add "end", CurCol - 1
    Result.Add "raw", Mid(Chars, Start, CurPos - Start)
    Result.Add "value", s
    Set ParseString = Result
End Function

Private Function ParseNumber() As Dictionary
    Dim Result As New Dictionary
    Dim Start As Long
    Start = CurPos
    Result.Add "line", CurLine
    Result.Add "start", CurCol
    Dim s As String
    Dim Dec As Boolean
    Dim Suffix As String
    Do While InString(DIGITS, CurChar) Or CurChar = "."
        If CurChar = "." And Dec Then
            ThrowError "Unexpected token '" & CurChar & "'.", CurLine, CurCol
        Else
            s = s & CurChar
            If CurChar = "." Then Dec = True
        End If
        If HasNextChar = False Then Exit Do
        NextChar
    Loop
    If Right(CurChar, 1) = "." Then ThrowError "Unexpected end of number."
    If InString("lfd", LCase(CurChar)) Then
        Suffix = CurChar
        NextChar
    End If
    Result.Add "end", CurCol - 1
    Result.Add "raw", Mid(Chars, Start, CurPos - Start)
    
    If Suffix = "" Then
        If InStr(1, s, ".") > 0 Then
            Result.Add "type", "double"
            Result.Add "value", CDbl(Replace(s, ".", Format(0, ".")))
        Else
            Result.Add "type", "integer"
            Result.Add "value", CLng(s)
        End If
    Else
        Select Case LCase(Suffix)
        Case "l"
            If InStr(1, s, ".") > 0 Then ThrowError "Syntax error."
            Result.Add "type", "long"
            Result.Add "value", CLngLng(s)
        Case "f"
            Result.Add "type", "float"
            Result.Add "value", CSng(Replace(s, ".", Format(0, ".")))
        Case "d"
            Result.Add "type", "double"
            Result.Add "value", CDbl(Replace(s, ".", Format(0, ".")))
        End Select
    End If
    
'    Dim Num As Variant
'
'    If InStr(1, s, ".") > 0 Then
'        Num = CDbl(Replace(s, ".", Format(0, ".")))
'        If (Num >= MIN_NEG_FLOAT Or Num >= MIN_POS_FLOAT) And (Num <= MAX_NEG_FLOAT Or Num <= MAX_POS_FLOAT) Then
'            Result.Add "type", "float"
'            Result.Add "value", CSng(Replace(s, ".", Format(0, ".")))
'        ElseIf (Num >= MIN_NEG_DOUBLE Or Num >= MIN_POS_DOUBLE) And (Num <= MAX_NEG_DOUBLE Or Num <= MAX_POS_DOUBLE) Then
'            Result.Add "type", "double"
'            Result.Add "value", CDbl(Replace(s, ".", Format(0, ".")))
'        Else
'            ThrowError "Overflow.", CurLine, Result("start")
'        End If
'    Else
'        #If Win64 Then
'            Num = CLngLng(s)
'        #Else
'            Num = CLng(s)
'        #End If
'        If Num >= MIN_HWORD And Num <= MAX_HWORD Then
'            Result.Add "type", "short"
'            Result.Add "value", CInt(s)
'        ElseIf Num >= MIN_WORD And Num <= MAX_WORD Then
'            Result.Add "type", "integer"
'            Result.Add "value", CLng(s)
'        #If Win64 Then
'        ElseIf Num >= MIN_DWORD And Num <= MAX_DWORD Then
'            Result.Add "type", "long"
'            Result.Add "value", CLngLng(s)
'        #End If
'        Else
'            ThrowError "Overflow.", CurLine, Result("start")
'        End If
'    End If
    
    Set ParseNumber = Result
End Function

Private Function ParseIdentifier() As Dictionary
    Set ParseIdentifier = New Dictionary
    Dim Start As Long
    Start = CurPos
    ParseIdentifier.Add "line", CurLine
    ParseIdentifier.Add "start", CurCol
    Dim s As String
    Do While IsValidIdentifierChar(CurChar)
        s = s & CurChar
        If HasNextChar = False Then Exit Do
        NextChar
    Loop
    ParseIdentifier.Add "end", CurCol - 1
    ParseIdentifier.Add "raw", Mid(Chars, Start, CurPos - Start)
    ParseIdentifier.Add "type", IIf(InString(KEYWORDS, "_" & s & "_"), "keyword", "identifier")
    ParseIdentifier.Add "value", s
End Function

Private Function ParseSymbol() As Dictionary
    If CurChar = "/" Then
        NextChar
        If CurChar = "/" Then
            pCommentLevel = 1
            NextChar
            Exit Function
        ElseIf CurChar = "*" Then
            pCommentLevel = 2
            NextChar
            Exit Function
        Else
            PrevChar
        End If
    End If
    
    Dim Symbol As New Dictionary
    Dim Start As Long
    Start = CurPos
    Symbol.Add "line", CurLine
    Symbol.Add "start", CurCol
    Dim s As String
    s = CurChar
    NextChar
    
    If InString("=!+-*/%<>|&^", s) And CurChar = "=" Then
        s = s & CurChar
        NextChar
    Else
        If InString("+-*<>|&", s) And CurChar = s Then
            s = s & CurChar
            NextChar
        End If
        
        If s = ">" And CurChar = ">" Then
            s = s & CurChar
            NextChar
        End If
        
        If s = ">>" And CurChar = ">" Then
            s = s & CurChar
            NextChar
        End If
        
        If (s = "**" Or s = "<<" Or s = ">>" Or s = ">>>" Or s = "&&" Or s = "||") And CurChar = "=" Then
            s = s & CurChar
            NextChar
        End If
    End If
    
'    If HasNextChar And InString(SYMBOLS, CurChar) Then
'        NextChar
'
'        If (s = "=" Or s = "<" Or s = ">" Or s = "+" Or s = "-" Or s = "*" Or s = "/" Or s = "!" Or s = "&" Or s = "|") And CurChar = "=" _
'            Or s = "+" And CurChar = "+" _
'            Or s = "-" And CurChar = "-" Then
'            s = s & CurChar
'            NextChar
'        End If
'    End If
    
    Symbol.Add "end", CurCol - 1
    Symbol.Add "raw", Mid(Chars, Start, CurPos - Start)
    Symbol.Add "type", "symbol"
    Symbol.Add "value", s
    Set ParseSymbol = Symbol
End Function

Private Function InString(String1 As String, String2 As String) As Boolean
    InString = InStr(1, String1, String2) > 0
End Function

Private Function IsValidIdentifierChar(Char As String) As Boolean
    IsValidIdentifierChar = InString(LETTERS, LCase(Char)) Or InString(DIGITS, Char) Or Char = "_"
End Function

Private Sub ThrowError(ErrMsg As String, Optional ByVal Line As Long = -1, Optional ByVal Col As Long = -1)
    Dim Msg As String
    Msg = "PARSING ERROR: " & ErrMsg & vbNewLine & "at line " & Line & " col " & Col & " at file '" & FileName & "'"
    If Line > 0 And Col > 0 Then Msg = Msg & vbNewLine & Lines(Line - 1) & vbNewLine & Space(Col - 1) & "^"
    RaiseEvent OnError(Msg)
    MsgBox Msg
    End
End Sub
