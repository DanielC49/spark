VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Spark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================
'
' ##############################################################
' #                                                            #
' #                       PPTGames Spark                       #
' #                    Programming Sparkuage                    #
' #                                                            #
' ##############################################################
'
' » version 0.1.0
'
' » https://pptgamespt.wixsite.com/pptg-coding/spark
'
' ===============================================================

Option Explicit

Public Event OnError(ErrMsg As String)
Public Event OnLog(LogMsg As String)

Private WithEvents Spark As SparkCore
Attribute Spark.VB_VarHelpID = -1
Private NativeLibraries As Dictionary
Private LibsToLoad As New Dictionary

Private Sub Class_Initialize()
    Set NativeLibraries = New Dictionary
    NativeLibraries.Add "io", True
    NativeLibraries.Add "math", True
    NativeLibraries.Add "string", True
    NativeLibraries.Add "sys", True
    NativeLibraries.Add "ppt", True
End Sub

Public Function CompileFile(File As String) As Double
    Dim t0 As Double
    t0 = Timer
    Set Spark = New SparkCore
    Set Spark.NativeLibraries = NativeLibraries
    
    Dim i As Long
    For i = 0 To LibsToLoad.Count - 1
        Spark.LoadLibrary LibsToLoad.Keys(i), LibsToLoad.Items(i)
    Next
    
    Spark.CompileFile File
    CompileFile = Timer - t0
End Function

Public Function Run(Optional FuncName As String = "main") As Double
    If Spark Is Nothing Then Err.Raise 1, , "Cannot run code." & vbNewLine & "Did you compile the code before?"
    Dim t0 As Double
    t0 = Timer
    Spark.RunFunction FuncName, New List
    Run = Timer - t0
End Function

Public Function GetTokens() As List
    Set GetTokens = Spark.Tokens
End Function

Public Function GetAst() As Dictionary
    Set GetAst = Spark.Ast
End Function

Public Sub SetNativeLibrary(LibId As String, Enabled As Boolean)
    If NativeLibraries.Exists(LibId) Then
        NativeLibraries(LibId) = Enabled
    Else
        Err.Raise 1, , "'" & LibId & "' is not a native library."
    End If
End Sub

Public Sub LoadLibrary(LibId As String, LibClass As Object)
    LibsToLoad.Add LibId, LibClass
End Sub

Private Sub Spark_OnError(ErrMsg As String)
    RaiseEvent OnError(ErrMsg)
End Sub

Private Sub Spark_OnLog(Msg As String)
    RaiseEvent OnLog(Msg)
End Sub
