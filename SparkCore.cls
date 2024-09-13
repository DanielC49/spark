VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SparkCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' short
'Private Const MIN_HWORD As Integer = -2 ^ 15
'Private Const MAX_HWORD As Integer = 2 ^ 15 - 1
'' integer
'Private Const MIN_WORD As Long = -2 ^ 31
'Private Const MAX_WORD As Long = 2 ^ 31 - 1
'' long
'Private Const MIN_DWORD As LongLong = -2 ^ 63
'Private Const MAX_DWORD As LongLong = 2 ^ 63 - 1
'
'' float
'Private Const MIN_POS_FLOAT As Single = 2 ^ -149
'Private Const MAX_POS_FLOAT As Single = (1 - 2 ^ -24) * 2 ^ 128
'Private Const MIN_NEG_FLOAT As Single = -(1 - 2 ^ -24) * 2 ^ 128
'Private Const MAX_NEG_FLOAT As Single = -2 ^ -149
'' double
'Private Const MIN_POS_DOUBLE As Double = 2 ^ -1074
'Private Const MAX_POS_DOUBLE As Double = (1 + (1 - 2 ^ -52)) * 2 ^ 1023
'Private Const MIN_NEG_DOUBLE As Double = -(1 + (1 - 2 ^ -52)) * 2 ^ 1023
'Private Const MAX_NEG_DOUBLE As Double = -2 ^ -1074

'Public Lines As List
Public Tokens As List
Public Ast As Dictionary

Public Libraries As New Dictionary
Public NativeLibraries As Dictionary

Private SparkNativeLibs As New SparkNativeLibs

Public FileName As String

Public ExportedFunctions As List
Public ExportedVariables As List

Public WithEvents Lexer As SparkLexer
Public WithEvents Parser As SparkParser
Public WithEvents Interpreter As SparkInterpreter

Public Event OnError(ErrMsg As String)
Public Event OnLog(Msg As String)

Public Sub CompileFile(File As String)
    File = Replace(File, "\", "/")
    FileName = Mid(File, InStrRev(File, "/"))
    
    If Dir(File) = "" Then ThrowError "File '" & File & "' not found."
    
    Dim Code As String
    Code = ReadTextFile(File)
    
    Set Lexer = New SparkLexer
    Set Parser = New SparkParser
    Set Interpreter = New SparkInterpreter
    
    Lexer.FileName = FileName
    Parser.FileName = FileName
    
    Set SparkNativeLibs.Spark = Me
    
    If NativeLibraries("io") Then LoadLibrary "io", SparkNativeLibs, True
    If NativeLibraries("math") Then LoadLibrary "math", SparkNativeLibs, True
    If NativeLibraries("string") Then LoadLibrary "string", SparkNativeLibs, True
    If NativeLibraries("sys") Then LoadLibrary "sys", SparkNativeLibs, True
    If NativeLibraries("ppt") Then LoadLibrary "ppt", SparkNativeLibs, True
    
    Set Interpreter.Spark = Me
    
    Set Tokens = Lexer.GetTokens(Code)
    Set Parser.Lines = Lexer.Lines
    Set Ast = Parser.Parse(Tokens)
    Set Interpreter.Libraries = Libraries
    Interpreter.Load Ast
End Sub

Public Function RunFunction(Name As String, Args As List) As Dictionary
    Dim FunctionCall As New Dictionary
    FunctionCall.Add "kind", "FunctionCall"
    FunctionCall.Add "name", Name
    FunctionCall.Add "arguments", Args
    Set RunFunction = Interpreter.RunFunction(FunctionCall)
End Function

Public Sub LoadLibrary(ByVal LibId As String, ByVal LibClass As Object, Optional IsNative As Boolean)
    Dim Lib As New Dictionary
    Dim LibLexer As New SparkLexer
    Dim LibParser As New SparkParser
    Dim LibTokens As List
    
    LibLexer.FileName = "Library: " & LibId
    LibParser.FileName = "Library: " & LibId
    
    If IsNative Then
        Set LibTokens = LibLexer.GetTokens(LibClass.GetDefinitions(LibId))
    Else
        Set LibTokens = LibLexer.GetTokens(LibClass.GetDefinitions)
    End If
    Set LibParser.Lines = LibLexer.Lines
    Libraries.Add LibId, LibParser.ParseLib(LibTokens, LibClass)
End Sub

Public Function EvalToString(ByVal Expression As Dictionary)
    EvalToString = Interpreter.EvaluateToString(Expression, Interpreter.CurFuncCallLine, Interpreter.CurFuncCallCol)
End Function

Public Function GetTypeKind(ByVal ValueType As String) As String
    Select Case ValueType
    Case "byte", "short", "integer", "long"
        GetTypeKind = "integer"
    Case "float", "double"
        GetTypeKind = "float"
    Case Else
        GetTypeKind = ValueType
    End Select
End Function

Public Sub Throw(ErrorMessage As String)
    ThrowError vbNewLine & ErrorMessage, Interpreter.CurFuncCallLine, Interpreter.CurFuncCallCol
End Sub

Public Sub LogMsg(Msg As String)
    RaiseEvent OnLog(Msg)
End Sub

Private Sub ThrowError(ErrMsg As String, Optional ByVal Line As Long = -1, Optional ByVal Col As Long = -1)
'    Err.Raise 1, , ErrMsg & vbNewLine & "at line " & Line & " col " & Col
    Dim Msg As String
    Msg = "ERROR: " & ErrMsg & vbNewLine & "at line " & Line & " col " & Col & " at file '" & FileName & "'"
    If Line > 0 And Col > 0 Then
        Msg = Msg & vbNewLine & Lexer.Lines(Line - 1) & vbNewLine & Space(Col - 1) & "^"
    End If
    
'    Err.Raise 1, , Msg
    RaiseEvent OnError(Msg)
    
    End
End Sub

Private Sub Lexer_OnError(Msg As String)
    RaiseEvent OnError(Msg)
End Sub

Private Sub Parser_OnError(Msg As String)
    RaiseEvent OnError(Msg)
End Sub

Private Sub Interpreter_OnError(Msg As String)
    RaiseEvent OnError(Msg)
End Sub

