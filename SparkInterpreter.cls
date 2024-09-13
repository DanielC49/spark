VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SparkInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Or Win64 Then
    Private Declare PtrSafe Function rtcCallByName Lib "VBE7.DLL" (ByVal Object As Object, ByVal ProcName As LongPtr, ByVal CallType As VbCallType, _
    ByRef Args() As Any, Optional ByVal lcid As Long) As Variant
#Else
    Private Declare Function rtcCallByName Lib "VBE6.DLL" (ByVal Object As Object, ByVal ProcName As Long, ByVal CallType As VbCallType, _
    ByRef args() As Any, Optional ByVal lcid As Long) As Variant
#End If

Public Spark As SparkCore

Private Ast As Dictionary

Private pGlobalContext As Dictionary
Private pCallStack As List

Private pFunctions As Dictionary
Private pGlobals As List
Private pLoadedLibraries As List
Private pLoadedFunctions As Dictionary

Public Libraries As Dictionary

'Private SparkLibraries As SparkLibs

Public CurFuncCallLine As Long
Public CurFuncCallCol As Long

Private NumTypesHi As Dictionary

Public Event OnError(ErrMsg As String)

Private Sub Class_Initialize()
    Set NumTypesHi = New Dictionary
    NumTypesHi.Add "byte", 0
    NumTypesHi.Add "short", 1
    NumTypesHi.Add "integer", 2
    NumTypesHi.Add "long", 3
    NumTypesHi.Add "float", 4
    NumTypesHi.Add "double", 5
End Sub

Public Sub Load(dAst As Dictionary)
    Set Ast = dAst
    
    Set pGlobals = Ast("globals")
    Set pFunctions = Ast("functions")
    
    Set pCallStack = New List
    Set pGlobalContext = New Dictionary
'    Set Libraries = New Dictionary
    
    pGlobalContext.Add "variables", New Dictionary
    pGlobalContext.Add "scope", 0
    
'    Set SparkLibraries = New SparkLibs
'    Set SparkLibraries.Spark = Spark
    
    LoadLibraries
    
'    Slide1.ListBox1.Clear
'    For i = 0 To pFunctions.Count - 1
'        Slide1.ListBox1.AddItem pFunctions.Keys(i)
'    Next
    
    Dim i As Long
    For i = 0 To pGlobals.Length - 1
        EvaluateVariableDeclaration pGlobals(i)
    Next
    
'    If DebugMode Then DebugStepWait = True
    
'    UpdateMemDisp
End Sub

Public Function RunFunction(FunctionCall As Dictionary) As Dictionary
    Set RunFunction = EvaluateFunctionCall(FunctionCall)
End Function

Private Sub LoadLibraries()
'    If Slide1.CheckBox2.Value Then Slide1.tbLibs.Text = JSON.Stringify(Libraries)
    
    Set pLoadedLibraries = New List
    Set pLoadedFunctions = New Dictionary
    
    Dim i As Long
    Dim Imports As List
    Set Imports = Ast("imports")
    For i = 0 To Imports.Length - 1
        EvaluateImportStatement Imports(i)
    Next
    
'    Slide1.tbLibs.Text = JSON.Stringify(Imports)
End Sub

Private Sub EvaluateImportStatement(Statement As Dictionary)
    Dim i As Long
    If Statement("type") = "library" Then
        Dim LibName As String
        LibName = Statement("value")
        If Libraries.Exists(LibName) Then
            If pLoadedLibraries.Includes(LibName) Then ThrowError "Duplicate import statement.", Statement("line"), Statement("start")
            pLoadedLibraries.Add LibName
            Dim LibFuncs As Dictionary
            Dim LibConsts As Dictionary
            Dim LibConst As Dictionary
            Set LibFuncs = Libraries(LibName)("functions")
            Set LibConsts = Libraries(LibName)("constants")
            For i = 0 To LibFuncs.Count - 1
                pLoadedFunctions.Add LibFuncs.Keys(i), LibFuncs.Items(i)
            Next
            For i = 0 To LibConsts.Count - 1
'                pGlobalContext("variables").Add LibConsts.Keys(i), LibConsts.Items(i)
                Set LibConst = LibConsts.Items(i)
                AddVariableToStack LibConst("name"), LibConst("type"), LibConst("dim"), True, pGlobalContext, Statement("line"), Statement("start")
                SetVariable LibConst, LibConst("value"), Statement("line"), Statement("start"), pGlobalContext
            Next
        Else
            ThrowError "Cannot import '" & LibName & "'.", Statement("line"), Statement("start")
        End If
    ElseIf Statement("type") = "file" Then
        Dim FilePath As String
        FilePath = Statement("value")
        If Dir(FilePath) = "" Then ThrowError "File '" & FilePath & "' not found.", Statement("line"), Statement("start")
        Dim Spark As New SparkCore
        Dim Funcs As List
        Dim Func As Dictionary
'        Set Spark.NativeLibraries = NativeLibraries
        Spark.CompileFile FilePath
        Set Funcs = Spark.ExportedFunctions
        For i = 0 To Funcs.Length - 1
            Set Func = New Dictionary
            Func.Add "kind", "ExportedFunction"
            Func.Add "name", Funcs(i)
            Func.Add "runtime", Spark
            Func.Add "file", FilePath
            pLoadedFunctions.Add Funcs(i), Func
        Next
        Dim Vars As List
        Set Vars = Spark.ExportedVariables
        For i = 0 To Vars.Length - 1
            pGlobals.Add Vars(i)
        Next
    End If
End Sub

Private Function EvaluateFunctionCall(FunctionCall As Dictionary) As Dictionary
    CurFuncCallLine = FunctionCall("line")
    CurFuncCallCol = FunctionCall("start")
    
    If pLoadedFunctions.Exists(FunctionCall("name")) Then
'        Set EvaluateFunctionCall = EvaluateNativeLibraryFunctionCall(FunctionCall)
        Dim Func As Dictionary
        Set Func = pLoadedFunctions(FunctionCall("name"))
        If Func("kind") = "NativeFunction" Then
            Set EvaluateFunctionCall = CallFunction(FunctionCall, Func)
        ElseIf Func("kind") = "ExportedFunction" Then
            Dim Args As List
            Dim EvalArgs As New List
            Set Args = FunctionCall("arguments")
            Dim i As Long
            For i = 0 To Args.Length - 1
                EvalArgs.Add EvaluateExpression(Args(i))
            Next
            Set EvaluateFunctionCall = Func("runtime").RunFunction(Func("name"), EvalArgs)
        Else
            ThrowError "FuncKind '" & Func("kind") & "' unknown"
        End If
    ElseIf FunctionCall("name") = "len" Then
        Dim Arr As Dictionary
        Set Arr = EvaluateExpression(FunctionCall("arguments")(0))
        If Arr("kind") = "ArrayExpression" Then
            Set EvaluateFunctionCall = NewLiteral("integer", UBound(Arr("value")) + 1)
        Else
            ThrowError "Expected array.", FunctionCall("arguments")(0)("line"), FunctionCall("arguments")(0)("start")
        End If
    ElseIf FunctionCall("name") = "copy" Then
        Set EvaluateFunctionCall = CallFuncCopy(FunctionCall)
    ElseIf pFunctions.Exists(FunctionCall("name")) Then
        Set EvaluateFunctionCall = CallFunction(FunctionCall, pFunctions(FunctionCall("name")))
    Else
        ThrowError "Function '" & FunctionCall("name") & "' not defined.", FunctionCall("line"), FunctionCall("start")
    End If
End Function

Private Function CallFuncCopy(FunctionCall As Dictionary) As Dictionary
    Dim Arr As Dictionary
    Set Arr = EvaluateExpression(FunctionCall("arguments")(0))
    Set CallFuncCopy = CopyArrExpr(Arr)
End Function

Private Function CopyArrExpr(Arr As Dictionary) As Dictionary
    Dim NewArr() As Variant
    ReDim NewArr(UBound(Arr("value")))
    Dim i As Long
    Dim Item As Dictionary
    
    For i = 0 To UBound(Arr("value"))
        Set Item = Arr("value")(i)
        If Item("kind") = "ArrayExpression" Then
            Set NewArr(i) = CopyArrExpr(Item)
        ElseIf Item("kind") = "Literal" Then
            Set NewArr(i) = NewLiteral(Item("type"), Item("value"))
        Else
            ThrowError "Cannot copy '" & Item("kind") & "'."
        End If
    Next
    
    Dim Result As New Dictionary
    Result.Add "kind", "ArrayExpression"
    Result.Add "type", Arr("type")
    Result.Add "dim", Arr("dim")
    Result.Add "value", NewArr
    Set CopyArrExpr = Result
End Function

Private Function CallFunction(FunctionCall As Dictionary, Func As Dictionary) As Dictionary
'    Dim Func As Dictionary
    Dim i As Long
    Dim Statements As List
'    Set Func = pFunctions(FunctionCall("name"))
    
    Dim IsNativeFunc As Boolean
    IsNativeFunc = Func("kind") = "NativeFunction"
    
    Dim FuncContext As New Dictionary
    If IsNativeFunc = False Then
        FuncContext.Add "name", FunctionCall("name")
        FuncContext.Add "line", FunctionCall("line")
        FuncContext.Add "start", FunctionCall("start")
        FuncContext.Add "variables", New Dictionary
        FuncContext.Add "current_scope", 0
    End If
    
    Dim Parameters As List
    Dim Args As List
    Dim Variable As Dictionary
    Dim Argument As Dictionary
    Dim Rest As Boolean
    Dim ReqParams As Long
    Set Parameters = Func("params")
    Set Args = FunctionCall("arguments")
    ReqParams = Func("requiredParams")
    
    If Args.Length < ReqParams Or Args.Length > Parameters.Length Then
        If Parameters.Length = 0 Then
            Rest = False
        Else
            Rest = Parameters.Last("rest")
        End If
        If Rest = False Or Args.Length < ReqParams - 1 Then
            If Parameters.Length = ReqParams Then
                ThrowError FunctionDefinitionToString(Func) & vbNewLine & _
                    "Expected " & ReqParams & " arguments but instead got " & Args.Length & ".", FunctionCall("line"), FunctionCall("start")
            Else
                ThrowError FunctionDefinitionToString(Func) & vbNewLine & _
                    "Expected between " & ReqParams & " and " & Parameters.Length & " arguments but instead got " & Args.Length & ".", FunctionCall("line"), FunctionCall("start")
            End If
        End If
    End If
    
    Dim Value As Variant
    Dim Args2() As Variant
    
    If Parameters.Length > 0 Then
        
        If IsNativeFunc Then ReDim Args2(Parameters.Length - 1) As Variant
        
        For i = 0 To Parameters.Length - 1
            If Parameters(i)("rest") Then Exit For
            If IsNativeFunc Then
                If i > Args.Length - 1 Then
                    If Parameters(i)("init") Is Nothing = False Then
                        Set Argument = EvaluateExpression(Parameters(i)("init"))
                    Else
                        ThrowError "PARAM ERROR"
                    End If
                Else
                    Set Argument = EvaluateExpression(Args(i))
                End If
                If Parameters(i)("type") = "any" Then
                    Set Args2(i) = Argument
                Else
                    If i > Args.Length - 1 Then
                        Set Argument = CastValue(Argument, Parameters(i)("type"), Parameters(i)("dim"), 0, 0)
                    Else
                        Set Argument = CastValue(Argument, Parameters(i)("type"), Parameters(i)("dim"), Args(i)("line"), Args(i)("start"))
                    End If
                    If Argument("kind") = "ArrayExpression" Then
                        Dim j As Long
                        Dim Arr() As Variant
                        ReDim Arr(UBound(Argument("value"))) As Variant
                        For j = 0 To UBound(Argument("value"))
                            Arr(j) = Argument("value")(j)("value")
                        Next
                        Argument("value") = Arr
                    End If
                    Args2(i) = Argument("value")
                End If
            Else
                If i > Args.Length - 1 Then
                    If Parameters(i)("init") Is Nothing = False Then
                        Set Argument = EvaluateExpression(Parameters(i)("init"))
                    Else
                        ThrowError "PARAM ERROR"
                    End If
                Else
                    Set Argument = EvaluateExpression(Args(i))
                End If
                AddVariableToStack Parameters(i)("name"), Parameters(i)("type"), Parameters(i)("dim"), False, FuncContext, Parameters(i)("line"), Parameters(i)("start")
                SetVariable Parameters(i), Argument, Parameters(i)("line"), Parameters(i)("start"), FuncContext
            End If
        Next
        
        If Parameters.Last("rest") Then
            Dim ArrArgs() As Variant
            Dim LastArg As Dictionary
            Dim RestArg As Dictionary
            
            If Args.Length >= Parameters.Length Then
                Set LastArg = EvaluateExpression(Args.Last)
                If Args.Length = Parameters.Length And LastArg("kind") = "ArrayExpression" Then
                    If LastArg("dim") = Parameters.Last("dim") + 1 Then Set RestArg = CastValue(LastArg, Parameters.Last("type"), Parameters.Last("dim") + 1, Args(Parameters.Length - 1)("line"), Args(Parameters.Length - 1)("start"))
                End If
                
                If RestArg Is Nothing Then
                    ReDim ArrArgs(Args.Length - Parameters.Length) As Variant
                    Dim Arg As Dictionary
                    Dim ArgEval As Dictionary
                    Dim Item As Dictionary
                    For i = 0 To Args.Length - Parameters.Length
                        Set Arg = Args(i + Parameters.Length - 1)
                        'If LastArg Is Nothing Then
                        If i = Args.Length - Parameters.Length Then
                            Set ArgEval = LastArg
                        Else
                            Set ArgEval = EvaluateExpression(Arg)
                        End If
                        Set Item = CastValue(ArgEval, Parameters.Last("type"), Parameters.Last("dim"), Arg("line"), Arg("start"))
                        Set ArrArgs(i) = Item
                    Next
                End If
            Else
                ReDim ArrArgs(0) As Variant
            End If
            
            If RestArg Is Nothing Then Set RestArg = MakeArrayExpression(Parameters.Last("type"), ArrArgs, Parameters.Last("dim") + 1)
            
            If IsNativeFunc Then
                Args2(UBound(Args2)) = RestArg("value")
            Else
                AddVariableToStack Parameters.Last("name"), Parameters.Last("type"), Parameters.Last("dim") + 1, False, FuncContext, Parameters.Last("line"), Parameters.Last("start")
                SetVariable Parameters.Last, RestArg, Parameters.Last("line"), Parameters.Last("start"), FuncContext
            End If
        End If
        
        If IsNativeFunc Then Value = CallByName2(Func("libclass"), "sparklib_" & FunctionCall("name"), VbMethod, Args2)
        
    Else
        If IsNativeFunc Then Value = CallByName(Func("libclass"), "sparklib_" & FunctionCall("name"), VbMethod)
    End If
    
    If IsNativeFunc = False Then pCallStack.Add FuncContext
    
    Dim Result As Dictionary
    If IsNativeFunc Then
        Set Result = LibValue(Value)
    Else
        Set Result = EvaluateFunctionBody(Func("body"))
        pCallStack.Remove pCallStack.Length - 1
    End If
    
    If Func("type") <> "void" Then
        If Result Is Nothing = False Then
            Set Result = CastValue(Result, Func("type"), Func("dim"), FunctionCall("line"), FunctionCall("start"))
        Else
            ThrowError "Function didn't return a value.", FunctionCall("line"), FunctionCall("start")
        End If
    ElseIf Func("type") = "void" And Result Is Nothing = False Then
        ' todo: improve loc
        ThrowError "Unexpected return value.", FunctionCall("line"), FunctionCall("start")
    End If
    
    Set CallFunction = Result
End Function

Private Function FunctionDefinitionToString(FunctionDefinition As Dictionary) As String
    Dim Result As String
    Dim i As Long
    Dim Params As List
    
    Result = FunctionDefinition("type") & RepeatString("[]", FunctionDefinition("dim")) & " " & FunctionDefinition("name") & "("
    
    Set Params = FunctionDefinition("params")
    For i = 0 To Params.Length - 1
        Result = Result & IIf(i = 0, "", ", ") & Params(i)("type") & RepeatString("[]", Params(i)("dim")) & " " & Params(i)("name")
        If Params(i)("init") Is Nothing = False Then
            Result = Result & " = " & IIf(Params(i)("init")("type") = "string", """" & Params(i)("init")("value") & """", Params(i)("init")("value"))
        End If
    Next
    
    Result = Result & ")"
    
    FunctionDefinitionToString = Result
End Function

Private Function EvaluateNativeLibraryFunctionCall(FunctionCall As Dictionary) As Dictionary
    Dim Func As Dictionary
    Dim i As Long
    Dim Args As List
    Set Func = pLoadedFunctions(FunctionCall("name"))
    Set Args = FunctionCall("arguments")
    
    If Args.Length <> UBound(Func("params")) + 1 Then ThrowError "Expected " & (UBound(Func("params")) + 1) & " arguments but instead got " & Args.Length & ".", FunctionCall("line"), FunctionCall("start")
    
    Dim Args2() As Variant
    Dim Arg As Dictionary
    Dim Value As Variant
    If Args.Length > 0 Then
        ReDim Args2(Args.Length - 1) As Variant
        For i = 0 To Args.Length - 1
            Set Arg = EvaluateExpression(Args(i))
            If Func("params")(i) = "any" Then
                Set Args2(i) = Arg
            Else
                Set Arg = CastValue(Arg, Func("params")(i), 1, Args(i)("line"), Args(i)("start"))
                Args2(i) = Arg("value")
            End If
        Next
        Value = CallByName2(Me, "sparklib_" & FunctionCall("name"), VbMethod, Args2)
    Else
        Value = CallByName(Me, "sparklib_" & FunctionCall("name"), VbMethod)
    End If
    Set EvaluateNativeLibraryFunctionCall = LibValue(Value)
End Function

Private Function LibValue(Value As Variant) As Dictionary
    If IsEmpty(Value) Then Exit Function
    Dim Result As New Dictionary
    If IsArray(Value) Then
        Result.Add "kind", "ArrayExpression"
        Result.Add "type", LibValueType(Value(0))
        Dim i As Long
        Dim Arr() As Variant
        ReDim Arr(UBound(Value)) As Variant
        For i = 0 To UBound(Value)
            Set Arr(i) = LibValue(Value(i))
        Next
        Result.Add "dim", 1
        Result.Add "value", Arr
    Else
        Result.Add "kind", "Literal"
        Result.Add "type", LibValueType(Value)
        Result.Add "value", Value
    End If
    Set LibValue = Result
End Function

Private Function LibValueType(Value As Variant) As String
    Select Case TypeName(Value)
    Case "Integer": LibValueType = "short"
    Case "Long": LibValueType = "integer"
    Case "LongLong": LibValueType = "long"
    Case "Single": LibValueType = "float"
    Case Else: LibValueType = LCase(TypeName(Value))
    End Select
End Function

Private Function EvaluateFunctionBody(Statements As List) As Dictionary
    Dim i As Long
    For i = 0 To Statements.Length - 1
'        If Statements(i)("kind") = "ReturnStatement" Then
'            Set EvaluateFunctionBody = EvaluateExpression(Statements(i)("value"))
'            Exit Function
'        Else
        If EvaluateStatement(Statements(i)) = "return" Or pCallStack.Last("returned") Then
            Set EvaluateFunctionBody = pCallStack.Last("return_value")
            Exit For
        End If
'        End If
    Next
End Function

Private Function GetVariable(ByVal Name As String, Optional Stack As Dictionary) As Dictionary
    If Stack Is Nothing Then Set Stack = pCallStack.Last
    If pGlobalContext("variables").Exists(Name) Then
        Set GetVariable = pGlobalContext("variables")(Name)
    ElseIf Stack("variables").Exists(Name) Then
        Set GetVariable = Stack("variables")(Name)
    Else
        ThrowError "Variable '" & Name & "' not defined."
    End If
End Function

Private Function GetVariableValue(ByVal Name As String, Line As Long, Start As Long) As Dictionary
    Dim Variable As Dictionary
    Set Variable = GetVariable(Name)("value")
    If Variable Is Nothing Then ThrowError "Variable '" & Name & "' not initialized.", Line, Start
    Set GetVariableValue = Variable
End Function

Private Sub SetVariable(ByVal Var As Dictionary, ByVal Value As Dictionary, ByVal Line As Long, ByVal Start As Long, Optional Stack As Dictionary)
    Dim Variable As Dictionary
    Dim VariableValue As Dictionary
    
    If Var("kind") = "Identifier" Or Var("kind") = "VariableDeclarator" Then
        Set Variable = GetVariable(Var("name"), Stack)
        If Variable("const") And Variable("value") Is Nothing = False Then ThrowError "Cannot assign to constant.", Line, Start
    ElseIf Var("kind") = "ArrayAccessExpression" Then
        Set VariableValue = EvaluateArrayAccessExpression2(Var)
'        If VariableValue("const") And VariableValue("value") Is Nothing = False Then ThrowError "Can't assign to constant.", Line, Start
    Else
        ThrowError "Syntax error.", Var("line"), Var("start")
    End If
    
    If Var("kind") = "ArrayAccessExpression" Then
        Set Value = CastValue(Value, VariableValue("type"), VariableValue("dim"), Line, Start)
'        If VariableValue("dim") <> Value("dim") Then ThrowError "Cannot cast from '" & Value("type") & RepeatString("[]", Value("dim")) _
            & "' to '" & VariableValue("type") & RepeatString("[]", VariableValue("dim")) & "'."
        VariableValue("value") = Value("value")
        VariableValue("type") = Value("type")
    Else
        Set Value = CastValue(Value, Variable("type"), Variable("dim"), Line, Start)
        If Value("kind") = "ArrayExpression" Then
'            If Variable("dim") <> Value("dim") Then ThrowError "Cannot cast from '" & Value("type") & RepeatString("[]", Value("dim")) _
                & "' to '" & Variable("type") & RepeatString("[]", Variable("dim")) & "'."
        End If
        Set Variable("value") = Value
    End If
End Sub

Private Function RepeatString(s As String, ByVal Count As Integer) As String
    Dim Result As String
    Dim i As Integer
    For i = 1 To Count
        Result = Result & s
    Next
    RepeatString = Result
End Function

Private Function CastValue(ByVal Value As Dictionary, ByVal CastType As String, ByVal ArrDim As Integer, ByVal Line As Long, ByVal Start As Long) As Dictionary
    Dim Result As Dictionary
    Set Result = Value
    If Value("kind") = "Literal" Then 'ThrowError "Cannot cast non-literal value." ' Exit Function
        If ArrDim > 0 Then ThrowError "Cannot cast from literal '" & Value("type") & RepeatString("[]", Value("dim")) _
            & "' to '" & CastType & RepeatString("[]", ArrDim) & "'.", Line, Start
        If Value("type") <> CastType Then
            If CastType <> "any" Then
                Select Case Value("type")
                Case "short", "integer", "long", "float", "double"
                    Dim v As Variant
                    v = Value("value")
                    Select Case CastType
                    Case "short"
                        Result("value") = CInt(Int(v))
                    Case "integer"
                        Result("value") = CLng(Int(v))
                    Case "long"
                        #If Win64 Then
                            Result("value") = CLngLng(Int(v))
                        #Else
                            ThrowError "Unsupported type.", Line, Start
                        #End If
                    Case "float"
                        Result("value") = CSng(v)
                    Case "double"
                        Result("value") = CDbl(v)
                    Case Else
                        ThrowError "Cannot cast from '" & Value("type") & "' to '" & CastType & "'.", Line, Start
                    End Select
                    Result("type") = CastType
                Case Else
                    If Value("type") <> CastType Then ThrowError "Cannot cast from '" & Value("type") & "' to '" & CastType & "'.", Line, Start
                End Select
            End If
        End If
    ElseIf Value("kind") = "ArrayExpression" Then
        If Value("dim") <> ArrDim And CastType <> "any" Then ThrowError "Cannot cast from array '" & Value("type") & RepeatString("[]", Value("dim")) _
            & "' to '" & CastType & RepeatString("[]", ArrDim) & "'.", Line, Start
        Dim i As Long
        Dim Arr As Variant
        Arr = Value("value")
        For i = 0 To UBound(Arr)
            Set Arr(i) = CastValue(Arr(i), CastType, ArrDim - 1, Line, Start)
        Next
        Result("value") = Arr
    End If
    Set CastValue = Result
End Function

Private Function DoesVariableExist(Name As String, Optional Stack As Dictionary) As Boolean
    If Stack Is Nothing Then Set Stack = pCallStack.Last
    DoesVariableExist = pGlobalContext("variables").Exists(Name) Or Stack("variables").Exists(Name)
End Function

Private Function EvaluateStatements(ByVal Statements As List) As String
    pCallStack.Last("current_scope") = pCallStack.Last("current_scope") + 1
    Dim i As Long
    Dim Result As String
    For i = 0 To Statements.Length - 1
        Result = EvaluateStatement(Statements(i))
        If Result <> "" Then
            EvaluateStatements = Result
            Exit For
        End If
    Next
    Dim Variables As New Dictionary
    For i = 0 To pCallStack.Last("variables").Count - 1
        If pCallStack.Last("variables").Items()(i)("scope") >= pCallStack.Last("current_scope") Then Exit For
        Variables.Add pCallStack.Last("variables").Keys()(i), pCallStack.Last("variables").Items()(i)
    Next
    Set pCallStack.Last("variables") = Variables
    pCallStack.Last("current_scope") = pCallStack.Last("current_scope") - 1
End Function

Private Sub AddVariable(ByVal Name As String, ByVal VariableType As String, ByVal VariableDim As Long, VariableConst As Boolean, ByVal Line As Long, ByVal Start As Long)
    AddVariableToStack Name, VariableType, VariableDim, VariableConst, pCallStack.Last, Line, Start
End Sub

Private Sub AddVariableToStack(ByVal Name As String, ByVal VariableType As String, ByVal VariableDim As Long, VariableConst As Boolean, Stack As Dictionary, ByVal Line As Long, ByVal Start As Long)
    If DoesVariableExist(Name, Stack) Then ThrowError "Duplicate declaration of variable '" & Name & "'.", Line, Start
    Dim Variable As New Dictionary
    Variable.Add "type", VariableType
    Variable.Add "dim", VariableDim
    Variable.Add "const", VariableConst
    Variable.Add "value", Nothing
    Variable.Add "scope", Stack("current_scope")
    Stack("variables").Add Name, Variable
End Sub

Private Function EvaluateStatement(Statement As Dictionary) As String
'    Do While DebugStepWait
'        DoEvents
'    Loop
'    DebugCurrentStatement = Statement("kind")
'    DebugCurrentLine = Statement("line")
'    Dim i As Long
'    Dim s As String
'    For i = 0 To UBound(Lines)
'        s = s & IIf(i = DebugCurrentLine - 1, ">", " ")
'        s = s & Lines(i) & vbNewLine
'    Next
'    DebugRunningCode = s
    
    Select Case Statement("kind")
    Case "VariableDeclaration"
        EvaluateVariableDeclaration Statement
    Case "ReturnStatement"
        EvaluateReturnStatement Statement
        EvaluateStatement = "return"
    Case "BreakStatement"
        EvaluateStatement = "break"
    Case "ContinueStatement"
        EvaluateStatement = "continue"
    Case "AssignmentExpression"
        EvaluateAssignmentExpression Statement
    Case "UpdateExpression"
        EvaluateUpdateExpression Statement
    Case "UnaryExpression"
        EvaluateUnaryExpression Statement
    Case "FunctionCall"
        EvaluateFunctionCall Statement
    Case "IfStatement"
        EvaluateStatement = EvaluateIfStatement(Statement)
    Case "SwitchStatement"
        EvaluateStatement = EvaluateSwitchStatement(Statement)
    Case "WhileLoop"
        EvaluateStatement = EvaluateWhileLoop(Statement)
    Case "DoWhileLoop"
        EvaluateStatement = EvaluateDoWhileLoop(Statement)
    Case "ForLoop"
        EvaluateStatement = EvaluateForLoop(Statement)
    Case "ForEachLoop"
        EvaluateStatement = EvaluateForEachLoop(Statement)
    Case "ThrowStatement"
        ThrowError EvaluateToString(EvaluateExpression(Statement("value")), Statement("line"), Statement("start")), Statement("line"), Statement("start")
    Case Else
        ThrowError "Unexpected token.", Statement("line"), Statement("start")
'        ThrowError "Unexpected statement '" & Statement("kind") & "'.", Statement("line"), Statement("start")
'        ThrowError "Unexpected statement '" & Statement("kind") & "'."
    End Select
'    DebugStepWait = True
End Function

Private Sub EvaluateVariableDeclaration(Statement As Dictionary)
    Dim i As Long
    For i = 0 To Statement("declarations").Length - 1
        EvaluateVariableDeclarator Statement("declarations")(i), Statement("const")
    Next
End Sub

Private Sub EvaluateVariableDeclarator(ByVal Statement As Dictionary, IsConst As Boolean)
'    If DoesVariableExist(Statement("name")) = False Then
        Dim InitValue As Dictionary
        Dim Stack As Dictionary
        Set InitValue = EvaluateExpression(Statement("init"))
        If Statement("global") Then
            Set Stack = pGlobalContext
        Else
            Set Stack = pCallStack.Last
        End If
        AddVariableToStack Statement("name"), Statement("type"), Statement("dim"), IsConst, Stack, Statement("line"), Statement("start")
        If InitValue Is Nothing = False Then SetVariable Statement, InitValue, Statement("line"), Statement("start"), Stack
'            UpdateMemDisp
'    Else
'        ThrowError "Duplicate declaration of variable '" & Statement("name") & "'."
'    End If
End Sub

Private Function EvaluateWhileLoop(Statement As Dictionary) As String
    Dim Statements As List
    Dim Result As String
    Do While EvaluateToBoolean(EvaluateExpression(Statement("condition")), Statement("condition")("line"), Statement("condition")("start"))
        Result = EvaluateStatements(Statement("statements"))
        If Result = "return" Then
            EvaluateWhileLoop = "return"
            Exit Function
        ElseIf Result = "break" Then
            Exit Do
        End If
    Loop
End Function

Private Function EvaluateDoWhileLoop(Statement As Dictionary) As String
    Dim Result As String
    Result = EvaluateStatements(Statement("statements"))
    If Result = "return" Or Result = "break" Then
        EvaluateDoWhileLoop = "return"
        Exit Function
    End If
    EvaluateDoWhileLoop = EvaluateWhileLoop(Statement)
End Function

Private Sub EvaluateReturnStatement(Statement As Dictionary)
    pCallStack.Last("returned") = True
    Set pCallStack.Last("return_value") = EvaluateExpression(Statement("value"))
End Sub

Private Function EvaluateForLoop(Statement As Dictionary) As String
    EvaluateStatement Statement("init")
    Dim Value As Dictionary
    Dim Result As String
    Set Value = EvaluateExpression(Statement("test"))
    Do While EvaluateToBoolean(Value, Statement("test")("line"), Statement("test")("start"))
        Result = EvaluateStatements(Statement("statements"))
        If Result = "return" Then
            EvaluateForLoop = "return"
            Exit Function
        ElseIf Result = "break" Then
            Exit Do
        End If
        EvaluateStatement Statement("update")
        Set Value = EvaluateExpression(Statement("test"))
    Loop
End Function

Private Function EvaluateForEachLoop(Statement As Dictionary) As String
    Dim Var As Dictionary
    Dim Arr As Dictionary
    Select Case Statement("left")("kind")
    Case "VariableDeclaration"
        EvaluateStatement Statement("left")
        Set Var = Statement("left")("declarations")(0)
    Case "Identifier"
        Set Var = Statement("left")
    Case Else
        ThrowError "Syntax error.", Statement("left")("line"), Statement("left")("start")
    End Select
    Set Arr = EvaluateExpression(Statement("right"))
    If Arr("kind") <> "ArrayExpression" Then ThrowError "Expected array.", Statement("right")("line"), Statement("right")("start")
    Dim i As Long
    Dim Result As String
    For i = 0 To UBound(Arr("value"))
        SetVariable Var, EvaluateExpression(Arr("value")(i)), -1, -1
        Result = EvaluateStatements(Statement("statements"))
        If Result = "return" Then
            EvaluateForEachLoop = "return"
            Exit Function
        ElseIf Result = "break" Then
            Exit For
        End If
    Next
End Function

Private Function EvaluateIfStatement(Statement As Dictionary) As String
    Dim Value As Dictionary
    Dim Result As String
    Set Value = EvaluateExpression(Statement("condition"))
    If EvaluateToBoolean(Value, Statement("condition")("line"), Statement("condition")("start")) Then
        Result = EvaluateStatements(Statement("consequent"))
    ElseIf Statement.Exists("alternate") Then
        Result = EvaluateStatements(Statement("alternate"))
    End If
    EvaluateIfStatement = Result
End Function

Private Function EvaluateSwitchStatement(Statement As Dictionary) As String
    Dim Discriminant As Dictionary
    Set Discriminant = EvaluateExpression(Statement("discriminant"))
    Dim i As Long
    Dim Cases As List
    Dim Value As Dictionary
    Dim Result As String
    Dim Passed As Boolean
    Set Cases = Statement("cases")
    For i = 0 To Cases.Length - 1
        If Cases(i)("test") Is Nothing Then
            Passed = True
        ElseIf Passed = False Then
            Set Value = EvaluateExpression(Cases(i)("test"))
            Passed = Spark.GetTypeKind(Discriminant("type")) = Spark.GetTypeKind(Value("type")) And Discriminant("value") = Value("value")
        End If
        If Passed Then
            Result = EvaluateStatements(Cases(i)("consequent"))
            If Result = "return" Then
                EvaluateSwitchStatement = "return"
                Exit Function
            ElseIf Result = "break" Then
                Exit For
            End If
        End If
    Next
End Function

Private Function EvaluateExpression(ByVal Expression As Dictionary) As Dictionary
    If Expression Is Nothing Then Exit Function
    Dim Result As Dictionary
    Set Result = New Dictionary
    If Expression("kind") = "Literal" Then
        Result.Add "kind", "Literal"
        Result.Add "type", Expression("type")
        Result.Add "value", Expression("value")
    ElseIf Expression("kind") = "ArrayExpression" Then
        Set Result = EvaluateArrayExpression(Expression)
    ElseIf Expression("kind") = "ArrayInitializerExpression" Then
        Set Result = EvaluateArrayInitializerExpression(Expression)
    ElseIf Expression("kind") = "ArrayAccessExpression" Then
        Set Result = EvaluateExpression(EvaluateArrayAccessExpression(Expression))
    ElseIf Expression("kind") = "AssignmentExpression" Then
        Set Result = EvaluateAssignmentExpression(Expression)
    ElseIf Expression("kind") = "BinaryExpression" Then
        Result.Add "kind", "Literal"
        Dim Lhs As Dictionary
        Dim Rhs As Dictionary
        Dim Value As Dictionary
        Set Lhs = EvaluateExpression(Expression("left"))
        Set Rhs = EvaluateExpression(Expression("right"))
        Set Value = EvaluateOperation(Expression("operator"), Lhs, Rhs, Expression("line"), Expression("start"))
'        Select Case Expression("operator")
'        Case "+": Value = Lhs("value") + Rhs("value")
'        Case "-": Value = Lhs("value") - Rhs("value")
'        Case "*": Value = Lhs("value") * Rhs("value")
'        Case "/": Value = Lhs("value") / Rhs("value")
'        End Select
        Result.Add "type", Value("type")
        Result.Add "value", Value("value")
    ElseIf Expression("kind") = "ParenthesizedExpression" Then
        Set Result = EvaluateExpression(Expression("expression"))
    ElseIf Expression("kind") = "ConditionalExpression" Then
        Dim Cond As Boolean
        Cond = EvaluateToBoolean(EvaluateExpression(Expression("condition")), Expression("condition")("line"), Expression("condition")("start"))
        If Cond Then
            Set Result = EvaluateExpression(Expression("consequent"))
        Else
            Set Result = EvaluateExpression(Expression("alternate"))
        End If
    ElseIf Expression("kind") = "UpdateExpression" Then
        Set Result = EvaluateUpdateExpression(Expression)
    ElseIf Expression("kind") = "UnaryExpression" Then
        Set Result = EvaluateUnaryExpression(Expression)
    ElseIf Expression("kind") = "CastExpression" Then
        Set Result = CastValue(EvaluateExpression(Expression("value")), Expression("type"), 0, Expression("line"), Expression("start"))
    ElseIf Expression("kind") = "Identifier" Then
        If DoesVariableExist(Expression("name")) Then
'            Dim Variable As Dictionary
            Set Result = GetVariableValue(Expression("name"), Expression("line"), Expression("start"))
'            Result.Add "kind", Variable("kind")
'            Result.Add "type", Variable("type")
'            Result.Add "value", Variable("value")
        Else
            ThrowError "Variable '" & Expression("name") & "' not defined.", Expression("line"), Expression("start")
        End If
    ElseIf Expression("kind") = "FunctionCall" Then
        Set Result = Nothing
        Set Result = EvaluateFunctionCall(Expression)
        If Result Is Nothing Then ThrowError "Function did not return a value.", Expression("line"), Expression("start")
    Else
        ThrowError "Unexpected expression '" & Expression("kind") & "'."
    End If
    Set EvaluateExpression = Result
End Function

Private Function EvaluateArrayInitializerExpression(Expression As Dictionary) As Dictionary
'    Dim Result As Dictionary
'    Dim i As Long
'    Dim Sizes As List
'    Set Sizes = Expression("sizes")
'    Set Result = NewArrayExpression(Expression("type"), Sizes(i))
'    For i = Sizes.Length - 1 To 0 Step -1
'        Set Result = NewArrayExpression(Expression("type"), Sizes(i))
'    Next
''    Result.Add "dim", ArrDim
''    Result.Add "value", Value
    Dim Sizes() As Variant
    Sizes = Expression("sizes")
    Set EvaluateArrayInitializerExpression = NewArrayExpression(Expression("type"), Sizes, 0)
End Function

Private Function NewArrayExpression(sType As String, Sizes() As Variant, Offset As Long) As Dictionary
    Dim Result As New Dictionary
    Dim i As Long
    Dim Arr() As Variant
    Dim Size As Dictionary
    Set Size = EvaluateExpression(Sizes(Offset))
    If Spark.GetTypeKind(Size("type")) <> "integer" Then ThrowError "Size must be an integer.", Sizes(Offset)("line"), Sizes(Offset)("start")
    If Size("value") = 0 Then ThrowError "Size must be greater than 0.", Sizes(Offset)("line"), Sizes(Offset)("start")
    ReDim Arr(Size("value") - 1) As Variant
    For i = 0 To Size("value") - 1
        If Offset = UBound(Sizes) Then
            Set Arr(i) = NewLiteral(sType, GetDefaultTypeValue(sType))
        Else
            Set Arr(i) = NewArrayExpression(sType, Sizes, Offset + 1)
        End If
    Next
    Result.Add "kind", "ArrayExpression"
    Result.Add "type", sType
    Result.Add "dim", UBound(Sizes) + 1
    Result.Add "value", Arr
    Set NewArrayExpression = Result
End Function

Private Function MakeArrayExpression(ByVal sType As String, Items() As Variant, ArrDim As Integer) As Dictionary
    Dim Result As New Dictionary
    Result.Add "kind", "ArrayExpression"
    Result.Add "type", sType
    Result.Add "dim", ArrDim
    Result.Add "value", Items
    Set MakeArrayExpression = Result
End Function

Private Function GetDefaultTypeValue(ValueType As String) As Variant
    Select Case ValueType
    Case "byte", "short", "integer", "long"
        GetDefaultTypeValue = 0
    Case "boolean"
        GetDefaultTypeValue = ""
    Case "string"
        GetDefaultTypeValue = ""
    Case Else
        ThrowError "Cannot generate default value for type '" & ValueType & "'."
    End Select
End Function

Private Function EvaluateArrayAccessExpression(Expression As Dictionary) As Dictionary
    Dim Obj As Dictionary
    Dim Index As Dictionary
    Dim Arr As Dictionary
    Set Obj = EvaluateExpression(Expression("object"))
    Set Index = EvaluateExpression(Expression("index"))
    
    If Obj("kind") <> "ArrayExpression" And Obj("kind") <> "Literal" And Obj("type") <> "string" Then ThrowError "Expected array or string.", Expression("object")("line"), Expression("object")("start")
    If Index("kind") <> "Literal" And Spark.GetTypeKind(Index("type")) <> "integer" Then ThrowError "Expected integer for index."
    
    Dim Result As Dictionary
    If Obj("kind") = "ArrayExpression" Then
        If Index("value") < 0 Or Index("value") > UBound(Obj("value")) Then ThrowError "Index out of bounds of array.", Expression("index")("line"), Expression("index")("start")
        Set Result = Obj("value")(Index("value"))
    ElseIf Obj("kind") = "Literal" And Obj("type") = "string" Then
        If Index("value") < 0 Or Index("value") > Len(Obj("value")) - 1 Then ThrowError "Index out of bounds of string.", Expression("index")("line"), Expression("index")("start")
        Set Result = NewLiteral("string", Mid(Obj("value"), Index("value") + 1, 1))
    End If
    If Result("kind") = "ArrayExpression" Then Set Result = EvaluateExpression(Result)
    Set EvaluateArrayAccessExpression = Result
End Function

Private Function EvaluateArrayAccessExpression2(Expression As Dictionary) As Dictionary
    Dim Obj As Dictionary
    Dim EvalObj As Dictionary
    Dim Index As Dictionary
    Dim Arr As Dictionary
    Set Obj = Expression("object")
    Set EvalObj = EvaluateExpression(Obj)
    Set Index = EvaluateExpression(Expression("index"))
    
    If Obj("kind") <> "Identifier" And Obj("kind") <> "ArrayAccessExpression" Then ThrowError "Expected variable.", Expression("object")("line"), Expression("object")("start")
    If Index("kind") <> "Literal" And Spark.GetTypeKind(Index("type")) <> "integer" Then ThrowError "Expected integer for index."
    If Index("value") < 0 Or Index("value") > UBound(EvalObj("value")) Then ThrowError "Index out of bounds of array.", Expression("index")("line"), Expression("index")("start")
    
    Dim Result As Dictionary
    If Obj("kind") = "ArrayAccessExpression" Then
        Set Result = EvaluateArrayAccessExpression2(Obj)("value")(Index("value"))
    Else ' Identifier
        Set Result = GetVariable(Obj("name"))("value")("value")(Index("value"))
    End If
    Set EvaluateArrayAccessExpression2 = Result
End Function

'Private Function EvaluateArrayAccessExpression2(Expression As Dictionary) As Dictionary
'    Dim Obj As Dictionary
'    Dim EvalObj As Dictionary
'    Dim Index As Dictionary
'    Dim Arr As Dictionary
'    Set Obj = Expression("object")
'    Set EvalObj = EvaluateExpression(Obj)
'    Set Index = EvaluateExpression(Expression("index"))
'
'    If Obj("kind") = "ArrayAccessExpression" Then Set Obj = EvaluateArrayAccessExpression2(Obj)
'
''    If Obj("kind") <> "ArrayExpression" Then ThrowError "Expected array.", Expression("object")("line"), Expression("object")("start")
'    If Obj("kind") <> "Identifier" Then ThrowError "Expected variable.", Expression("object")("line"), Expression("object")("start")
'    If Index("kind") <> "Literal" And Index("type") <> "integer" Then ThrowError "Expected integer for index."
'    If Index("value") < 0 Or Index("value") > UBound(EvalObj("value")) Then ThrowError "Index out of bounds of array.", Expression("index")("line"), Expression("index")("start")
'
'    Dim Result As Dictionary
'    Set Result = GetVariable(Obj("name"))("value")("value")(Index("value"))
''    If Result("kind") = "ArrayExpression" Then Set Result = Result("value")
'    Set EvaluateArrayAccessExpression2 = Result
'End Function

Private Function EvaluateArrayExpression(Expression As Dictionary) As Dictionary
    Dim Result As New Dictionary
    Result.Add "kind", "ArrayExpression"
    Dim Arr() As Variant
    Dim Value() As Variant
    Dim ArrDim As Integer
    Dim Item As Dictionary
    Dim sType As String
    Dim i As Long
    Arr = Expression("value")
    ReDim Value(UBound(Arr))
    For i = 0 To UBound(Arr)
        Set Item = EvaluateExpression(Arr(i))
        If i = 0 Then
            sType = Item("type")
        ElseIf Item("type") <> sType Then
            ThrowError "Illegal value type in array at index " & i & ". Expected '" & sType & "'.", Arr(i)("line"), Arr(i)("start")
        End If
        If ArrDim = 0 Then
            If Item("kind") = "ArrayExpression" Then
                ArrDim = Item("dim") + 1
            Else
                ArrDim = 1
            End If
        Else
            If Item("kind") = "ArrayExpression" Then
                If Item("dim") + 1 <> ArrDim Then ThrowError "Invalid array."
            Else
                If ArrDim > 1 Then ThrowError "Invalid array."
            End If
        End If
        Set Value(i) = Item
    Next
    Result.Add "type", sType 'Expression("type")
    Result.Add "dim", ArrDim
    Result.Add "value", Value
    Set EvaluateArrayExpression = Result
End Function

Private Function EvaluateUpdateExpression(Expression As Dictionary) As Dictionary
    Dim Name As String
    Name = Expression("value")("name")
    If DoesVariableExist(Name) Then
'        Dim Result As New Dictionary
        Dim Value1 As Dictionary
        Dim Value2 As New Dictionary
        Set Value1 = GetVariableValue(Name, Expression("line"), Expression("start"))
        Value2.Add "kind", "Literal"
        Value2.Add "type", Value1("type")
        If Expression("operator") = "++" Then
            Value2.Add "value", Value1("value") + 1
        ElseIf Expression("operator") = "--" Then
            Value2.Add "value", Value1("value") - 1
        Else
            ThrowError "Invalid update operator '" & Expression("operator") & "'."
        End If
        SetVariable Expression("value"), Value2, Expression("line"), Expression("start")
        If Expression("prefix") Then
            Set EvaluateUpdateExpression = Value2
        Else
            Set EvaluateUpdateExpression = Value1
        End If
'        UpdateMemDisp
    Else
        ThrowError "Variable '" & Name & "' not defined.", Expression("line"), Expression("start")
    End If
End Function

Private Function EvaluateUnaryExpression(Expression As Dictionary) As Dictionary
    Dim Result As New Dictionary
    Dim Value As Dictionary
    Set Value = EvaluateExpression(Expression("value"))
    Result.Add "kind", "Literal"
    Result.Add "type", Value("type")
    If Expression("operator") = "+" Then
        If IsNumber(Spark.GetTypeKind(Value("type"))) = False Then ThrowError "Bad operand type for operator '" & Expression("operator") & "'.", Expression("line"), Expression("start")
        Result.Add "value", Value("value")
    ElseIf Expression("operator") = "-" Then
        If IsNumber(Spark.GetTypeKind(Value("type"))) = False Then ThrowError "Bad operand type for operator '" & Expression("operator") & "'.", Expression("line"), Expression("start")
        Result.Add "value", -Value("value")
    ElseIf Expression("operator") = "!" Then
        Result.Add "value", Not EvaluateToBoolean(Value, Expression("value")("line"), Expression("value")("start"))
    ElseIf Expression("operator") = "typeof" Then
        Result("type") = "string"
        Result.Add "value", Value("type") & RepeatString("[]", Value("dim"))
    Else
        ThrowError "Unknown unary operator '" & Expression("operator") & "'."
    End If
    Set EvaluateUnaryExpression = Result
End Function

Private Function EvaluateAssignmentExpression(Expression As Dictionary) As Dictionary
    Dim Name As String
    If Expression("left")("kind") = "Identifier" Then
        Name = Expression("left")("name")
'    ElseIf Expression("left")("kind") = "ArrayAccessExpression" Then
'        Name = Expression("left")("object")("name")
    End If
    If DoesVariableExist(Name) Or Name = "" Then
        Dim Result As New Dictionary
        Dim Lhs As Dictionary
        Dim Rhs As Dictionary
        Dim Value As Dictionary
        Set Rhs = EvaluateExpression(Expression("right"))
        If Expression("operator") = "=" Then
            Set Value = Rhs
        Else
            Set Lhs = EvaluateExpression(Expression("left"))
            Set Value = EvaluateOperation(Mid(Expression("operator"), 1, Len(Expression("operator")) - 1), Lhs, Rhs, Expression("line"), Expression("start"))
        End If
        Set Result = Value
'        Result.Add "kind", Value("kind") '"Literal"
'        Result.Add "type", Value("type")
'        Result.Add "value", Value("value")
        SetVariable Expression("left"), Result, Expression("line"), Expression("start")
        Set EvaluateAssignmentExpression = Result
'        UpdateMemDisp
    Else
        ThrowError "Variable '" & Name & "' not defined.", Expression("line"), Expression("start")
    End If
End Function

Private Function MaxInt(a As Integer, B As Integer) As Integer
    If a > B Then
        MaxInt = a
    Else
        MaxInt = B
    End If
End Function

Private Function IsNumber(TypeKind As String) As Boolean
    IsNumber = TypeKind = "integer" Or TypeKind = "float"
End Function

Private Function EvaluateOperation(Operator As String, Lhs As Dictionary, Rhs As Dictionary, Line As Long, Start As Long) As Dictionary
    Dim Result As New Dictionary
    Dim LhsTypeKind As String
    Dim RhsTypeKind As String
    LhsTypeKind = Spark.GetTypeKind(Lhs("type"))
    RhsTypeKind = Spark.GetTypeKind(Rhs("type"))
    If Operator = "+" Then
        If LhsTypeKind = "string" Or RhsTypeKind = "string" Then
            Result.Add "kind", "Literal"
            Result.Add "type", "string"
            Result.Add "value", CStr(Lhs("value")) & CStr(Rhs("value"))
        ElseIf IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            If LhsTypeKind = "integer" And RhsTypeKind = "integer" Then
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") + Rhs("value")
            ElseIf LhsTypeKind = "float" Or RhsTypeKind = "float" Then
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") + Rhs("value")
            End If
        End If
    ElseIf Operator = "-" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            If LhsTypeKind = "integer" And RhsTypeKind = "integer" Then
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") - Rhs("value")
            ElseIf LhsTypeKind = "float" Or RhsTypeKind = "float" Then
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") - Rhs("value")
            End If
        End If
    ElseIf Operator = "*" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            If LhsTypeKind = "integer" And RhsTypeKind = "integer" Then
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") * Rhs("value")
            ElseIf LhsTypeKind = "float" Or RhsTypeKind = "float" Then
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") * Rhs("value")
            End If
        End If
    ElseIf Operator = "/" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            If LhsTypeKind = "integer" And RhsTypeKind = "integer" Then
                If Rhs("value") = 0 Then ThrowError "Division by zero."
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", CLng(Lhs("value") / Rhs("value"))
            ElseIf LhsTypeKind = "float" Or RhsTypeKind = "float" Then
                If Rhs("value") = 0 Then ThrowError "Division by zero."
                Result.Add "kind", "Literal"
                Result.Add "type", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type"))))
                Result.Add "value", Lhs("value") / Rhs("value")
            End If
        End If
    ElseIf Operator = "%" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", Lhs("value") Mod Rhs("value")
        End If
    ElseIf Operator = "&" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", CLng(Lhs("value") And Rhs("value"))
        End If
    ElseIf Operator = "|" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", CLng(Lhs("value") Or Rhs("value"))
        End If
    ElseIf Operator = "^" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", CLng(Lhs("value") Xor Rhs("value"))
        End If
    ElseIf Operator = "&&" Then
        If LhsTypeKind = "boolean" And RhsTypeKind = "boolean" Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") And Rhs("value")
'        Else
'            Dim Value As Long
'            If EvaluateToBoolean(Lhs("value")) Then
'                Value = Rhs("value")
'            Else
'                Value = Lhs("value")
'            End If
'            Result.Add "value", Value
        End If
    ElseIf Operator = "||" Then
        If LhsTypeKind = "boolean" And RhsTypeKind = "boolean" Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") Or Rhs("value")
        Else
            ThrowError "Bad operands for binary operator '" & Operator & "'."
        End If
    ElseIf Operator = "^" Then
        If LhsTypeKind = "boolean" And RhsTypeKind = "boolean" Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") Xor Rhs("value")
        Else
            ThrowError "Bad operands for binary operator '" & Operator & "'."
        End If
    ElseIf Operator = "^^" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", CLng(Lhs("value") Xor Rhs("value"))
        End If
    ElseIf Operator = "<<" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            If Rhs("value") > 0 Then
                Dim i As Long
                Dim m As Long
                Dim Shl As Long
                Shl = Lhs("value")
                For i = 1 To Rhs("value")
                    m = Shl And &H40000000
                    Shl = (Shl And &H3FFFFFFF) * 2
                    If m <> 0 Then
                        Shl = Shl Or &H80000000
                    End If
                Next
            End If
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", IIf(Rhs("value") > 0, Shl, 0)
        End If
    ElseIf Operator = "**" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", IIf(LhsTypeKind = "integer" And RhsTypeKind = "integer", "integer", NumTypesHi.Keys(MaxInt(NumTypesHi(Lhs("type")), NumTypesHi(Rhs("type")))))
            Result.Add "value", Lhs("value") ^ Rhs("value")
        End If
    ElseIf Operator = ">>" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", IIf(Rhs("value") > 0, Int(Lhs("value") / (2 ^ Rhs("value"))), Lhs("value"))
        End If
    ElseIf Operator = ">>>" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "integer"
            Result.Add "value", 0
        End If
    ElseIf Operator = "==" Then
        Result.Add "kind", "Literal"
        Result.Add "type", "boolean"
        Result.Add "value", Spark.GetTypeKind(Lhs("type")) = Spark.GetTypeKind(Rhs("type")) And Lhs("value") = Rhs("value")
    ElseIf Operator = "!=" Then
        Result.Add "kind", "Literal"
        Result.Add "type", "boolean"
        Result.Add "value", Lhs("type") = Rhs("type") Or Lhs("value") <> Rhs("value")
    ElseIf Operator = "<" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") < Rhs("value")
        End If
    ElseIf Operator = ">" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") > Rhs("value")
        End If
    ElseIf Operator = "<=" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") <= Rhs("value")
        End If
    ElseIf Operator = ">=" Then
        If IsNumber(LhsTypeKind) Or IsNumber(RhsTypeKind) Then
            Result.Add "kind", "Literal"
            Result.Add "type", "boolean"
            Result.Add "value", Lhs("value") >= Rhs("value")
        End If
    Else
        ThrowError "Invalid operator '" & Operator & "'.", Line, Start
    End If
    If Result.Count = 0 Then ThrowError "Bad operand types for operator '" & Operator & "'.", Line, Start
    Set EvaluateOperation = Result
End Function

Private Function NewLiteral(LiteralType As String, LiteralValue As Variant) As Dictionary
    Set NewLiteral = New Dictionary
    NewLiteral.Add "kind", "Literal"
    NewLiteral.Add "type", LiteralType
    NewLiteral.Add "value", LiteralValue
End Function

Private Function EvaluateToBoolean(Expression As Dictionary, ByVal Line As Long, ByVal Start As Long) As Boolean
    If Expression("kind") = "Literal" Then
        If Expression("type") = "boolean" Then
            EvaluateToBoolean = Expression("value")
        ElseIf Spark.GetTypeKind(Expression("type")) = "integer" Or Spark.GetTypeKind(Expression("type")) = "float" Then
            EvaluateToBoolean = Expression("value") <> 0
        ElseIf Expression("type") = "string" Then
            EvaluateToBoolean = Expression("value") <> ""
        Else
            ThrowError "Cannot '" & Expression("value") & "' evaluate to boolean.", Line, Start
        End If
    Else
        ThrowError "Cannot evaluate '" & Expression("kind") & "' to boolean.", Line, Start
    End If
End Function

Public Function EvaluateToString(ByVal Expression As Dictionary, Line As Long, Start As Long) As String
    If Expression Is Nothing Then
        ThrowError "Cannot cast 'void' to string.", Line, Start
    ElseIf Expression("kind") = "Literal" Then
        If Expression("type") = "string" Then
            EvaluateToString = Expression("value")
        ElseIf Spark.GetTypeKind(Expression("type")) = "integer" Or Spark.GetTypeKind(Expression("type")) = "float" Then
            EvaluateToString = Replace(CStr(Expression("value")), Format(0, "."), ".")
        ElseIf Expression("type") = "boolean" Then
            EvaluateToString = IIf(Expression("value"), "true", "false")
        Else
            ThrowError "Unexpected type '" & Expression("type") & "'."
        End If
    ElseIf Expression("kind") = "ArrayExpression" Then
        Dim i As Long
        Dim Arr() As Variant
        Dim s As String
        Arr = Expression("value")
        For i = 0 To UBound(Arr)
            s = s & EvaluateToString(Arr(i), Line, Start) & IIf(i < UBound(Arr), ",", "")
        Next
        EvaluateToString = "{" & s & "}"
    Else
        ThrowError "Cannot cast '" & Expression("kind") & "' to string.", Line, Start
    End If
End Function

Private Function CallByName2(Object As Object, ProcName As String, CallType As VbCallType, Args() As Variant) As Variant
   AssignResult CallByName2, rtcCallByName(Object, StrPtr(ProcName), CallType, Args)
End Function

Private Sub AssignResult(Target As Variant, Result As Variant)
  If VBA.IsObject(Result) Then Set Target = Result Else Target = Result
End Sub

Private Sub ThrowError(ErrMsg As String, Optional ByVal Line As Long = -1, Optional ByVal Col As Long = -1)
'    Err.Raise 1, , ErrMsg & vbNewLine & "at line " & Line & " col " & Col
    Dim Msg As String
    Msg = "ERROR: " & ErrMsg & vbNewLine & "at line " & Line & " col " & Col & " at file '" & Spark.FileName & "'"
    If Line > 0 And Col > 0 Then
        Msg = Msg & vbNewLine & Spark.Lexer.Lines(Line - 1) & vbNewLine & Space(Col - 1) & "^"
    End If
    
    If pCallStack Is Nothing = False Then
        Dim i As Long
        For i = pCallStack.Length - 1 To 0 Step -1
            Msg = Msg & vbNewLine & "  at '" & pCallStack(i)("name") & "'"
            If pCallStack(i).Exists("line") Then Msg = Msg & " (" & pCallStack(i)("line") & ":" & pCallStack(i)("start") & ")"
        Next
    End If
    
'    Err.Raise 1, , Msg
    RaiseEvent OnError(Msg)
    
    End
End Sub
