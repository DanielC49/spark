VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SparkParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileName As String
Public Lines As List

Private pTokens As List
Private CurTokenPos As Long
Private CurToken As Dictionary

Private pImports As List
Private pGlobals As List
Private pFunctions As Dictionary

Public ExportedFunctions As List
Public ExportedVariables As List

Private mOperators As New Dictionary
Private mDeclTypes As New Dictionary

Public Event OnError(ErrMsg As String)

Private Sub Class_Initialize()
    mOperators.RemoveAll
    ' operator, precedence
    mOperators.Add "=", 4
    mOperators.Add "+=", 4
    mOperators.Add "-=", 4
    mOperators.Add "*=", 4
    mOperators.Add "/=", 4
    mOperators.Add "**=", 4
    mOperators.Add "%=", 4
    mOperators.Add "<<=", 4
    mOperators.Add ">>=", 4
    mOperators.Add ">>>=", 4
    mOperators.Add "&=", 4
    mOperators.Add "^=", 4
    mOperators.Add "|=", 4
    mOperators.Add "&&=", 4
    mOperators.Add "||=", 4
    mOperators.Add "||", 5
    mOperators.Add "&&", 6
    mOperators.Add "|", 7
    mOperators.Add "^", 8
    mOperators.Add "&", 9
    mOperators.Add "==", 10
    mOperators.Add "!=", 10
    mOperators.Add "<", 11
    mOperators.Add ">", 11
    mOperators.Add "<=", 11
    mOperators.Add ">=", 11
    mOperators.Add "<<", 12
    mOperators.Add ">>", 12
    mOperators.Add ">>>", 12
    mOperators.Add "+", 13
    mOperators.Add "-", 13
    mOperators.Add "*", 14
    mOperators.Add "/", 14
    mOperators.Add "%", 14
    mOperators.Add "**", 15
    
    mDeclTypes.RemoveAll
    mDeclTypes.Add "void", "void"
    mDeclTypes.Add "any", "any"
    mDeclTypes.Add "byte", "byte"
    mDeclTypes.Add "short", "short"
    mDeclTypes.Add "int", "integer"
    mDeclTypes.Add "long", "long"
    mDeclTypes.Add "float", "float"
    mDeclTypes.Add "double", "double"
    mDeclTypes.Add "bool", "boolean"
    mDeclTypes.Add "string", "string"
End Sub

Public Function Parse(Tokens As List) As Dictionary
    Dim Ast As New Dictionary
    
    Set pTokens = Tokens
    CurTokenPos = -1
    NextToken
    
    Set pImports = New List
    Set pFunctions = New Dictionary
    Ast.Add "functions", pFunctions
    
    Set pGlobals = New List
    Ast.Add "globals", pGlobals
    
    Set ExportedFunctions = New List
    Set ExportedVariables = New List
    
    Do While CurTokenPos < pTokens.Length
        Dim Stmt As Dictionary
        Set Stmt = ExpectFunctionDefinition
        If Stmt Is Nothing Then Set Stmt = ExpectImportStatement
        If Stmt Is Nothing Then Set Stmt = ExpectGlobalVariableDeclaration
        If Stmt Is Nothing Then ThrowError "Unexpected token '" & CurToken("raw") & "'.", CurToken("line"), CurToken("start")
    Loop
    
    Ast.Add "imports", pImports
    
    Set Parse = Ast
End Function

Private Sub NextToken()
    CurTokenPos = CurTokenPos + 1
    If CurTokenPos < pTokens.Length Then Set CurToken = pTokens(CurTokenPos)
End Sub

Private Sub PrevToken()
    CurTokenPos = CurTokenPos - 1
    Set CurToken = pTokens(CurTokenPos)
End Sub

Private Sub SetCurToken(Pos As Long)
    CurTokenPos = Pos
    If CurTokenPos < pTokens.Length Then Set CurToken = pTokens(CurTokenPos)
End Sub

Private Function CompareToken(Token As Dictionary, TokenType As String, TokenValue As Variant) As Boolean
    CompareToken = Token("type") = TokenType And Token("value") = TokenValue
End Function

Private Function ExpectFunctionDefinition() As Dictionary
    Dim ParseStart As Long
    ParseStart = CurTokenPos
    
    Dim Modifier As String
    Modifier = ExpectModifier
    
    Dim PossibleType As String
    Dim ArrayDim As Integer
    PossibleType = ExpectType
    ArrayDim = ExpectVariableArray
    If PossibleType <> "" Then
        Dim PossibleName As Dictionary
        Set PossibleName = ExpectIdentifier
        If Not PossibleName Is Nothing Then
            Dim PossibleSymbol As Dictionary
            Set PossibleSymbol = ExpectSymbol("(")
            If Not PossibleSymbol Is Nothing Then
                Dim Func As New Dictionary
                Dim Params As New List
                Func.Add "kind", "FunctionDefinition"
                Func.Add "name", PossibleName("value")
                Func.Add "type", mDeclTypes(PossibleType)
                Func.Add "dim", ArrayDim
                Func.Add "public", Modifier = "public"
                
                Dim Param As Dictionary
                Dim ReqParams As Long
                Dim HasOptionalParam As Boolean
                
                Do While ExpectSymbol(")") Is Nothing
'                    PossibleType = ExpectType
'                    If PossibleType = "" Then ThrowError "Expected parameter type.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
'                    Dim PossibleVariableName As Dictionary
'                    Set PossibleVariableName = ExpectIdentifier
'                    If PossibleVariableName Is Nothing Then ThrowError "Expected parameter name.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
                    
'                    Set Param = New Dictionary
'                    Param.Add "type", PossibleType
'                    Param.Add "name", PossibleVariableName("value")
                    Dim Variable As Dictionary
                    Set Variable = ExpectVariableDeclaration(True)
                    If Variable Is Nothing Then ThrowError "Expected parameter type.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
'                    Param.Add "kind", "Identifier"
'                    Param.Add "type", Variable("type")
'                    Param.Add "name", Variable("name")
'                    Param.Add "dim", Variable("dim")
'                    Param.Add "init", Variable("init")
                    Variable("kind") = "Identifier"
                    If Variable("init") Is Nothing Then
                        ReqParams = ReqParams + 1
                        If HasOptionalParam Then ThrowError "Optional parameters may only appear after required parameters", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
                    Else
                        HasOptionalParam = True
                        Set Variable("init") = ValueConstant(Variable("init"))
                        If Variable("init")("kind") <> "Literal" Then ThrowError "Default parameter value must be constant.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
                        If Variable("dim") <> 0 Or Variable("rest") = True Then ThrowError "Default parameter cannot be array.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
                        If GetTypeKind(Variable("init")("type")) <> GetTypeKind(Variable("type")) Then ThrowError "Type mismatch.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
                    End If
                    Params.Add Variable
                    
                    If ExpectSymbol(")") Is Nothing = False Then
                        Exit Do
                    End If
                    
                    If ExpectSymbol(",") Is Nothing Then
                        ThrowError "Expected ',' or ')'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
                    ElseIf Variable("rest") Then
                        ThrowError "Comma not permitted after rest parameter.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
                    End If
                Loop
                
                Func.Add "params", Params
                Func.Add "requiredParams", ReqParams
                Func.Add "body", ParseFunctionBody
                
                pFunctions.Add Func("name"), Func
                If Func("public") Then ExportedFunctions.Add Func("name")
                
                Set ExpectFunctionDefinition = Func
            Else
                SetCurToken ParseStart
            End If
        Else
            SetCurToken ParseStart
        End If
    Else
        SetCurToken ParseStart
    End If
End Function

Private Function ExpectImportStatement() As Dictionary
    If ExpectKeyword("import") Is Nothing Then Exit Function
    
    Dim Result As New Dictionary
    
    Result.Add "kind", "ImportStatement"
    Result.Add "line", CurToken("line")
    Result.Add "start", CurToken("start")
    
    Dim LibName As Dictionary
    Set LibName = ExpectOneValue
    
    If LibName("kind") <> "Literal" Or LibName("type") <> "string" Then ThrowError "Expected string.", CurToken("line"), CurToken("start")
    
    Dim Name As String
    Name = LibName("value")
    
    If Left(Name, 2) = "./" Then
        Result.Add "type", "file"
        Name = ActivePresentation.Path & "\" & Mid(Name, 3)
    Else
        Result.Add "type", "library"
    End If
    
    Result.Add "value", Name
    
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';' at end of statement.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    pImports.Add Result
    
    Set ExpectImportStatement = Result
End Function

Private Function ParseFunctionBody() As List
    If ExpectSymbol("{") Is Nothing Then ThrowError "Expected '{'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    If CurTokenPos >= pTokens.Length Then ThrowError "Unexpected end of input.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Dim Statements As New List
    Dim Statement As Dictionary
    
    Do While ExpectSymbol("}") Is Nothing
        Set Statement = ExpectStatement
        If Statement Is Nothing = False Then
            Statements.Add Statement
        Else
            ThrowError "Unexpected token '" & CurToken("value") & "'.", CurToken("line"), CurToken("start")
        End If
        
'        If ExpectSymbol(";") Is Nothing Then
'            ThrowError "Expected ';' at end of statement."
'        End If
    Loop
    
    Set ParseFunctionBody = Statements
    
'    If ExpectSymbol("}") Is Nothing Then ThrowError "Unbalanced '{'."
End Function

Private Function ExpectStatementBlock() As List
    If ExpectSymbol("{") Is Nothing Then Exit Function 'ThrowError "Expected '{'."
    
    Dim Statements As New List
    Dim Statement As Dictionary
    
    Do While ExpectSymbol("}") Is Nothing
        Set Statement = ExpectStatement
        If Statement Is Nothing = False Then
            Statements.Add Statement
        Else
            ThrowError "Unexpected token '" & CurToken("value") & "'.", CurToken("line"), CurToken("start")
        End If
        
'        If ExpectSymbol(";") Is Nothing Then
'            ThrowError "Expected ';' at end of statement."
'        End If
    Loop
    
'    If ExpectSymbol("}") Is Nothing Then ThrowError "Unbalanced '{'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Set ExpectStatementBlock = Statements
End Function

Private Function ExpectStatement(Optional RequireEnd As Boolean = True) As Dictionary
    Dim Result As Dictionary
    Set Result = ExpectIfStatement
    If Result Is Nothing Then Set Result = ExpectSwitchStatement
    If Result Is Nothing Then Set Result = ExpectWhileLoop
    If Result Is Nothing Then Set Result = ExpectDoWhileLoop
    If Result Is Nothing Then Set Result = ExpectForLoop
    If Result Is Nothing Then Set Result = ExpectReturnStatement
    If Result Is Nothing Then Set Result = ExpectBreakStatement
    If Result Is Nothing Then Set Result = ExpectContinueStatement
    If Result Is Nothing Then Set Result = ExpectThrowStatement
    If Result Is Nothing Then
        Set Result = ExpectVariableDeclaration(False)
        If Result Is Nothing = False And RequireEnd Then If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';' at end of statement.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    End If
    If Result Is Nothing Then
        Set Result = ExpectExpression
        If Result Is Nothing = False And RequireEnd Then If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';' at end of statement.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    End If
    Set ExpectStatement = Result
End Function

Private Function ExpectThrowStatement() As Dictionary
    Dim Line As Long
    Dim Start As Long
    Line = CurToken("line")
    Start = CurToken("start")
    
    If ExpectKeyword("throw") Is Nothing Then Exit Function
    
    Dim Expr As Dictionary
    Set Expr = ExpectExpression
    
    If Expr Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';'."
    
    Dim Result As New Dictionary
    Result.Add "kind", "ThrowStatement"
    Result.Add "line", Line
    Result.Add "start", Start
    Result.Add "value", Expr
    
    Set ExpectThrowStatement = Result
End Function

Private Function ExpectReturnStatement() As Dictionary
    If ExpectKeyword("return") Is Nothing Then Exit Function
    
    Dim Result As New Dictionary
    Result.Add "kind", "ReturnStatement"
    Result.Add "line", CurToken("line")
    Result.Add "start", CurToken("start")
    
    Result.Add "value", ExpectExpression
    
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';'."
    
    Set ExpectReturnStatement = Result
End Function

Private Function ExpectBreakStatement() As Dictionary
    If ExpectKeyword("break") Is Nothing Then Exit Function
    
    Dim Result As New Dictionary
    Result.Add "kind", "BreakStatement"
    Result.Add "line", CurToken("line")
    
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';'."
    
    Set ExpectBreakStatement = Result
End Function

Private Function ExpectContinueStatement() As Dictionary
    If ExpectKeyword("continue") Is Nothing Then Exit Function
    
    Dim Result As New Dictionary
    Result.Add "kind", "ContinueStatement"
    Result.Add "line", CurToken("line")
    
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';'."
    
    Set ExpectContinueStatement = Result
End Function

Private Function ExpectVariableDeclaration(AsParameter As Boolean, Optional IsGlobal As Boolean) As Dictionary
    Dim IsConst As Boolean
    IsConst = ExpectKeyword("const") Is Nothing = False
    
    Dim PossibleType As String
    PossibleType = ExpectType
    If PossibleType = "" Or PossibleType = "void" Then
        If IsConst Then ThrowError "Expected type.", CurToken("line"), CurToken("start")
        If PossibleType = "void" Then PrevToken
        Exit Function
    End If
    
    Dim Declarations As New List
    Dim Declarator As Dictionary
    
    Do
        Set Declarator = ExpectVariableDeclarator(PossibleType, AsParameter, IsGlobal)
        If Declarator Is Nothing Then Exit Function
        Declarations.Add Declarator
        If AsParameter = False Then If ExpectSymbol(",") Is Nothing Then Exit Do
        If AsParameter Then Exit Do
    Loop
    
    Dim Statement As Dictionary
    
    If AsParameter Then
        Set Statement = Declarator
        Statement("kind") = "Identifier"
    Else
        Set Statement = New Dictionary
        Statement.Add "kind", "VariableDeclaration"
        Statement.Add "line", Declarations(0)("line")
        Statement.Add "start", Declarations(0)("start")
        Statement.Add "global", IsGlobal
        Statement.Add "const", IsConst
        Statement.Add "declarations", Declarations
    End If
    
    Set ExpectVariableDeclaration = Statement
End Function

Private Function ExpectGlobalVariableDeclaration() As Dictionary
    Dim Modifier As String
    Modifier = ExpectModifier
    
    Dim Result As Dictionary
    Set Result = ExpectVariableDeclaration(False, True)
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';' at end of statement.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Result.Add "public", Modifier = "public"
    
    pGlobals.Add Result
    If Result("public") Then ExportedVariables.Add Result
    
    Set ExpectGlobalVariableDeclaration = Result
End Function

Private Function ExpectVariableArray() As Integer
    Dim ArrayDim As Integer
    Do While ExpectSymbol("[") Is Nothing = False
        If ExpectSymbol("]") Is Nothing Then ThrowError "Expected ']'.", CurToken("line"), CurToken("start")
        ArrayDim = ArrayDim + 1
    Loop
    ExpectVariableArray = ArrayDim
End Function

Private Function ExpectVariableDeclarator(sType As String, AsParameter As Boolean, Optional IsGlobal As Boolean) As Dictionary
    Dim StartToken As Long
    StartToken = CurTokenPos
    
    Dim Start As Long
    Start = CurToken("start")
    
    Dim ArrayDim As Integer
    ArrayDim = ExpectVariableArray
    
    Dim AsRest As Boolean
    
    If AsParameter Then
        If ExpectSymbol(".") Is Nothing = False Then
            If ExpectSymbol(".") Is Nothing Then ThrowError "Unexpected token '" & pTokens(CurTokenPos - 1)("value") & "'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
            If ExpectSymbol(".") Is Nothing Then ThrowError "Unexpected token '" & pTokens(CurTokenPos - 1)("value") & "'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
            AsRest = True
        End If
    End If
    
    Dim PossibleVariableName As Dictionary
    Set PossibleVariableName = ExpectIdentifier
    If PossibleVariableName Is Nothing Then
        ThrowError "Expected variable name.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 2
'        SetCurToken StartToken
'        Exit Function
    End If
    
    Dim Statement As New Dictionary
    Statement.Add "kind", "VariableDeclarator"
    Statement.Add "line", pTokens(CurTokenPos - 2)("line")
    Statement.Add "start", Start
    Statement.Add "global", IsGlobal
    Statement.Add "type", mDeclTypes(sType)
    Statement.Add "dim", ArrayDim
    Statement.Add "name", PossibleVariableName("value")
    If AsParameter Then Statement.Add "rest", AsRest
    
'    If AsParameter = False Then
        Dim InitialValue As Dictionary
        
        If ExpectSymbol("=") Is Nothing = False Then
            Set InitialValue = ExpectExpression
            If InitialValue Is Nothing Then
                ThrowError "Expected expression for right-hand side assignment.", CurToken("line"), CurToken("start")
            ElseIf InitialValue("type") <> sType Then
    '            ThrowError "Type mismatch. Expected type '" & PossibleType & "'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
            End If
        Else
            Set InitialValue = Nothing
        End If
        
        Statement.Add "init", InitialValue
'    End If
    
    Set ExpectVariableDeclarator = Statement
End Function

Private Function ExpectFunctionCall() As Dictionary
    Dim StartToken As Long
    StartToken = CurTokenPos
    
    Dim PossibleFunctionName As Dictionary
    Set PossibleFunctionName = ExpectIdentifier
    
    If PossibleFunctionName Is Nothing Then
        SetCurToken StartToken
        Exit Function
    End If
    
    If ExpectSymbol("(") Is Nothing Then
        SetCurToken StartToken
        Exit Function
    End If
    
    Dim FunctionCall As New Dictionary
    FunctionCall.Add "kind", "FunctionCall"
    FunctionCall.Add "line", pTokens(CurTokenPos - 2)("line")
    FunctionCall.Add "start", pTokens(CurTokenPos - 2)("start")
    FunctionCall.Add "name", PossibleFunctionName("value")
    
    Dim Arguments As New List
    Dim Argument As Dictionary
    
    Do While ExpectSymbol(")") Is Nothing
        Set Argument = ExpectExpression
        If Argument Is Nothing Then
            ThrowError "Expected expression as argument.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
        End If
        
        Arguments.Add Argument
        
        If ExpectSymbol(")") Is Nothing = False Then
            Exit Do
        End If
        
        If ExpectSymbol(",") Is Nothing Then
            ThrowError "Expected ',' or ')'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
        End If
    Loop
    
    FunctionCall.Add "arguments", Arguments
    
'    If ExpectSymbol(";") Is Nothing Then
'        ThrowError "Expected ';' at end of function call."
'    End If
    
    Set ExpectFunctionCall = FunctionCall
End Function

Private Function ExpectWhileLoop() As Dictionary
    If ExpectKeyword("while") Is Nothing Then Exit Function
    If ExpectSymbol("(") Is Nothing Then ThrowError "Expected '('.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Dim Condition As Dictionary
    Set Condition = ExpectExpression
    If Condition Is Nothing Then ThrowError "Expected expression for condition.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Dim WhileLoop As New Dictionary
    WhileLoop.Add "kind", "WhileLoop"
    WhileLoop.Add "line", CurToken("line")
    WhileLoop.Add "condition", Condition
    
    If ExpectSymbol(")") Is Nothing Then ThrowError "Expected ')'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Dim Statements As List
    Set Statements = ExpectStatementOrBlock
    If Statements Is Nothing Then ThrowError "Expected statement or statement block.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    WhileLoop.Add "statements", Statements
    
    Set ExpectWhileLoop = WhileLoop
End Function

Private Function ExpectDoWhileLoop() As Dictionary
    If ExpectKeyword("do") Is Nothing Then Exit Function
    
    Dim DoWhileLoop As New Dictionary
    DoWhileLoop.Add "kind", "DoWhileLoop"
    DoWhileLoop.Add "line", CurToken("line")
    
    Dim Statements As List
    Set Statements = ExpectStatementOrBlock
    If Statements Is Nothing Then ThrowError "Expected statement or statement block.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    DoWhileLoop.Add "statements", Statements
    
    If ExpectKeyword("while") Is Nothing Then Exit Function
    
    If ExpectSymbol("(") Is Nothing Then ThrowError "Expected '('.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Dim Condition As Dictionary
    Set Condition = ExpectExpression
    If Condition Is Nothing Then ThrowError "Expected expression for condition.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    DoWhileLoop.Add "condition", Condition
    
    If ExpectSymbol(")") Is Nothing Then ThrowError "Expected ')'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';'.", CurToken("line"), CurToken("start")
    
    Set ExpectDoWhileLoop = DoWhileLoop
End Function

Private Function ExpectForLoop() As Dictionary
    If ExpectKeyword("for") Is Nothing Then Exit Function
    If ExpectSymbol("(") Is Nothing Then ThrowError "Expected '('."
    
    Dim Stmt As Dictionary
    Set Stmt = ExpectStatement(False)
    If Stmt Is Nothing Then ThrowError "Expected expression."
    
    Dim ForLoop As Dictionary
    
    If ExpectSymbol(":") Is Nothing Then
        Set ForLoop = ExpectNormalForLoop(Stmt)
    Else
        Set ForLoop = ExpectForEachLoop(Stmt)
    End If
    
    If ExpectSymbol(")") Is Nothing Then ThrowError "Expected ')'."
    
    Dim Statements As List
    Set Statements = ExpectStatementOrBlock
    If Statements Is Nothing Then ThrowError "Expected statement or statement block.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    ForLoop.Add "statements", Statements
    
    Set ExpectForLoop = ForLoop
End Function

Private Function ExpectNormalForLoop(InitExpr As Dictionary) As Dictionary
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';'.", CurToken("line"), CurToken("start")
    Dim TestExpr As Dictionary
    Dim UpdateExpr As Dictionary
    Set TestExpr = ExpectStatement
    Set UpdateExpr = ExpectStatement(False)
    If TestExpr Is Nothing Then ThrowError "Expected expression."
    If UpdateExpr Is Nothing Then ThrowError "Expected expression."
    
    Dim ForLoop As New Dictionary
    ForLoop.Add "kind", "ForLoop"
    ForLoop.Add "line", CurToken("line")
    ForLoop.Add "start", CurToken("start")
    ForLoop.Add "init", InitExpr
    ForLoop.Add "test", TestExpr
    ForLoop.Add "update", UpdateExpr
    
    Set ExpectNormalForLoop = ForLoop
End Function

Private Function ExpectForEachLoop(Lhs As Dictionary)
    Dim Rhs As Dictionary
    Set Rhs = ExpectStatement(False)
    
    If Rhs Is Nothing Then ThrowError "Expected expression."
    
    Dim ForLoop As New Dictionary
    ForLoop.Add "kind", "ForEachLoop"
    ForLoop.Add "line", CurToken("line")
    ForLoop.Add "start", CurToken("start")
    ForLoop.Add "left", Lhs
    ForLoop.Add "right", Rhs
    
    Set ExpectForEachLoop = ForLoop
End Function

Private Function ExpectIfStatement() As Dictionary
    If ExpectKeyword("if") Is Nothing Then Exit Function
    If ExpectSymbol("(") Is Nothing Then ThrowError "Expected '('."
    
    Dim Condition As Dictionary
    Set Condition = ExpectExpression
    If Condition Is Nothing Then ThrowError "Expected expression for condition."
    
    Dim IfStatement As New Dictionary
    IfStatement.Add "kind", "IfStatement"
    IfStatement.Add "line", CurToken("line")
    IfStatement.Add "condition", Condition
    
    If ExpectSymbol(")") Is Nothing Then ThrowError "Expected ')'."
    
    Dim Statements As List
    Set Statements = ExpectStatementOrBlock
    If Statements Is Nothing Then ThrowError "Expected statement or statement block.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    IfStatement.Add "consequent", Statements
    
    If ExpectKeyword("else") Is Nothing = False Then
        IfStatement.Add "alternate", ExpectStatementOrBlock
    End If
    
    Set ExpectIfStatement = IfStatement
End Function

Private Function ExpectSwitchStatement() As Dictionary
    If ExpectKeyword("switch") Is Nothing Then Exit Function
    If ExpectSymbol("(") Is Nothing Then ThrowError "Expected '('.", CurToken("line"), CurToken("start")
    
    Dim Discriminant As Dictionary
    Set Discriminant = ExpectExpression
    If Discriminant Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
    If ExpectSymbol(")") Is Nothing Then ThrowError "Expected ')'.", CurToken("line"), CurToken("start")
    If ExpectSymbol("{") Is Nothing Then ThrowError "Expected '{'.", CurToken("line"), CurToken("start")
    
    Dim Cases As New List
    Dim HasDefault As Boolean
    
    Do While ExpectSymbol("}") Is Nothing
        If CurToken("type") = "keyword" And CurToken("value") = "default" Then
            If HasDefault Then
                ThrowError "Multiple default clauses.", CurToken("line"), CurToken("start")
            Else
                HasDefault = True
            End If
        End If
        Cases.Add ExpectSwitchCase
    Loop
    
    Dim SwitchStatement As New Dictionary
    SwitchStatement.Add "kind", "SwitchStatement"
    SwitchStatement.Add "discriminant", Discriminant
    SwitchStatement.Add "cases", Cases
    
    Set ExpectSwitchStatement = SwitchStatement
End Function

Private Function ExpectSwitchCase() As Dictionary
    Dim IsDefault As Boolean
    If ExpectKeyword("case") Is Nothing Then
        If ExpectKeyword("default") Is Nothing Then ThrowError "Expected 'case' or 'default'."
        IsDefault = True
    End If
    
    If IsDefault = False Then
        Dim TestExpr As Dictionary
        Set TestExpr = ExpectExpression
        If TestExpr Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
    End If
    
    If ExpectSymbol(":") Is Nothing Then ThrowError "Expected ':'.", CurToken("line"), CurToken("start")
    
    Dim Consequent As New List
    
    Do While ExpectKeyword("case") Is Nothing And ExpectKeyword("default") Is Nothing And ExpectSymbol("}") Is Nothing
        Consequent.Add ExpectStatement
    Loop
    
    PrevToken
    
    Dim SwitchCase As New Dictionary
    SwitchCase.Add "kind", "SwitchCase"
    SwitchCase.Add "test", TestExpr
    SwitchCase.Add "consequent", Consequent
    
    Set ExpectSwitchCase = SwitchCase
End Function

Private Function ExpectStatementOrBlock() As List
    Dim Result As List
    Set Result = ExpectStatementBlock
    If Result Is Nothing Then
        Set Result = New List
        Dim Statement As Dictionary
        Set Statement = ExpectStatement
        If Statement Is Nothing Then Exit Function
        Result.Add Statement
    End If
    Set ExpectStatementOrBlock = Result
End Function

Private Function ExpectExpression() As Dictionary
    Dim PossibleArrayInitializer As String
    PossibleArrayInitializer = ExpectType
    
    If PossibleArrayInitializer <> "" And PossibleArrayInitializer <> "void" Then
        Set ExpectExpression = ExpectArrayInitializer(PossibleArrayInitializer)
        Exit Function
    ElseIf PossibleArrayInitializer <> "" Then
        PrevToken
    End If
    
    Dim Lhs As Dictionary
    Dim Rhs As Dictionary
    
    Set Lhs = ExpectOneValue
    If Lhs Is Nothing Then Exit Function
    
    Dim Operator As Dictionary
    Dim BinaryExpression As Dictionary
    Do
        Set Operator = ExpectSymbol
        If Operator Is Nothing Then Exit Do
        Dim RhsPrecedence As Long
        RhsPrecedence = GetOperatorPrec(Operator("value"))
        If RhsPrecedence = 0 Then
            If Operator("value") = "?" Then
                Set ExpectExpression = ExpectConditionalExpression(Lhs)
                Exit Function
            Else
                PrevToken
                Set ExpectExpression = Lhs
                Exit Function
            End If
        End If
        Set Rhs = ExpectOneValue
        If Rhs Is Nothing Then
            ThrowError "Expected expression", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
'            PrevToken
'            Set ExpectExpression = Lhs
'            Exit Function
        End If
        
        Dim RightmostStatement As Dictionary
        Set RightmostStatement = FindRightmostStatement(Lhs, RhsPrecedence)
        If RightmostStatement Is Nothing = False Then
            Set BinaryExpression = ParseOperation(Operator("value"), RightmostStatement("right"), Rhs)
            BinaryExpression.Add "start", Operator("start")
            Set RightmostStatement("right") = BinaryExpression
        Else
            Set BinaryExpression = ParseOperation(Operator("value"), Lhs, Rhs)
            BinaryExpression.Add "start", Operator("start")
            Set Lhs = BinaryExpression
        End If
    Loop
    
    Set ExpectExpression = Lhs
End Function

Private Function ExpectArrayInitializer(sType As String) As Dictionary
    Dim ArraySize As Dictionary
    Dim ArraySizes As New List
    
    If ExpectSymbol("[") Is Nothing Then ThrowError "Expected '['", CurToken("line"), CurToken("start")
    PrevToken
    
    Do While ExpectSymbol("[") Is Nothing = False
        Set ArraySize = ExpectExpression
        If ArraySize Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
        If ExpectSymbol("]") Is Nothing Then ThrowError "Expected ']'.", CurToken("line"), CurToken("start")
        ArraySizes.Add ArraySize
    Loop
    
    Dim Result As New Dictionary
    Result.Add "kind", "ArrayInitializerExpression"
    Result.Add "type", mDeclTypes(sType)
    Result.Add "line", CurToken("line")
    Result.Add "start", CurToken("start")
    Result.Add "dim", ArraySizes.Length
    Result.Add "sizes", ArraySizes.Items
    
    Set ExpectArrayInitializer = Result
End Function

Private Function ExpectConditionalExpression(Condition As Dictionary) As Dictionary
    Dim Result As New Dictionary
    Result.Add "kind", "ConditionalExpression"
    Result.Add "condition", Condition
    
    Dim Consequent As Dictionary
    Dim Alternate As Dictionary
    
    Set Consequent = ExpectExpression
    If Consequent Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
    If ExpectSymbol(":") Is Nothing Then ThrowError "Expected ':'.", CurToken("line"), CurToken("start")
    
    Set Alternate = ExpectExpression
    If Alternate Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
    
    Result.Add "consequent", Consequent
    Result.Add "alternate", Alternate
    
    Set ExpectConditionalExpression = Result
End Function

Private Function ParseOperation(Operator As String, Lhs As Dictionary, Rhs As Dictionary)
    Dim Operation As New Dictionary
    If Right(Operator, 1) = "=" And Operator <> "==" And Operator <> "!=" And Operator <> "<=" And Operator <> ">=" Then
        Operation.Add "kind", "AssignmentExpression"
        Operation.Add "line", CurToken("line")
    Else
        Operation.Add "kind", "BinaryExpression"
        Operation.Add "line", CurToken("line")
    End If
    Operation.Add "operator", Operator
    Operation.Add "left", Lhs
    Operation.Add "right", Rhs
    Set ParseOperation = Operation
End Function

Private Function FindRightmostStatement(Lhs As Dictionary, RhsPrecedence As Long) As Dictionary
    If Lhs("kind") <> "AssignmentExpression" And Lhs("kind") <> "BinaryExpression" Then Exit Function
    If GetOperatorPrec(Lhs("operator")) >= RhsPrecedence Then Exit Function
    
    Dim Rhs As Dictionary
    Set Rhs = Lhs("right")
    Set Rhs = FindRightmostStatement(Rhs, RhsPrecedence)
    If Rhs Is Nothing Then
        Set FindRightmostStatement = Lhs
    Else
        Set FindRightmostStatement = Rhs
    End If
End Function

Private Function GetOperatorPrec(Symbol As String) As Long
    If mOperators.Exists(Symbol) Then GetOperatorPrec = mOperators(Symbol)
'    IsOperator = InString("_+__-__*__/_", "_" & Symbol & "_")
End Function

Private Function ExpectOneValue() As Dictionary
    Dim Result As Dictionary
    If GetTypeKind(CurToken("type")) = "integer" Or GetTypeKind(CurToken("type")) = "float" Or CurToken("type") = "string" Then
        Set Result = New Dictionary
        Result.Add "kind", "Literal"
        Result.Add "line", CurToken("line")
        Result.Add "start", CurToken("start")
        Result.Add "type", CurToken("type")
        Result.Add "value", CurToken("value")
        NextToken
    ElseIf CurToken("type") = "keyword" And (CurToken("value") = "false" Or CurToken("value") = "true") Then
        Set Result = New Dictionary
        Result.Add "kind", "Literal"
        Result.Add "line", CurToken("line")
        Result.Add "start", CurToken("start")
        Result.Add "type", "boolean"
        Result.Add "value", CurToken("value") = "true"
        NextToken
    ElseIf CurToken("type") = "identifier" Then
        Set Result = ExpectFunctionCall
        If Result Is Nothing Then
            Dim Value As New Dictionary
            Value.Add "kind", "Identifier"
            Value.Add "name", CurToken("value")
            Value.Add "line", CurToken("line")
            Value.Add "start", CurToken("start")
            Value.Add "end", CurToken("end")
            NextToken
            If CurToken("type") = "symbol" Then
                If CurToken("value") = "++" Or CurToken("value") = "--" Then
                    Set Result = New Dictionary
                    Result.Add "kind", "UpdateExpression"
                    Result.Add "operator", CurToken("value")
                    Result.Add "prefix", False
                    Result.Add "value", Value
                    NextToken
                Else
                    Set Result = Value
                End If
            Else
                Set Result = Value
            End If
        End If
    ElseIf ExpectSymbol("(") Is Nothing = False Then
        Set Result = New Dictionary
        Dim PossibleType As String
        PossibleType = ExpectType(False)
        If PossibleType <> "" Then
            If ExpectSymbol(")") Is Nothing Then ThrowError "Expected '('.", CurToken("line"), CurToken("start")
            Dim v As Dictionary
            Set v = ExpectOneValue()
            If v Is Nothing Then ThrowError "Expected value.", CurToken("line"), CurToken("start")
            Result.Add "kind", "CastExpression"
            Result.Add "type", mDeclTypes(PossibleType)
            Result.Add "value", v
        Else
            Dim Expr As Dictionary
            Set Expr = ExpectExpression
            If Expr Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
            Result.Add "kind", "ParenthesizedExpression"
            Result.Add "expression", Expr
            If ExpectSymbol(")") Is Nothing Then ThrowError "Unbalanced '('.", CurToken("line"), CurToken("start")
        End If
    ElseIf CurToken("type") = "keyword" And CurToken("value") = "typeof" Then
        NextToken
        Dim Value3 As Dictionary
        Set Value3 = ExpectOneValue()
        If Value3 Is Nothing Then ThrowError "Expected value.", CurToken("line"), CurToken("start")
        Set Result = New Dictionary
        Result.Add "kind", "UnaryExpression"
        Result.Add "line", CurToken("line")
        Result.Add "start", CurToken("start")
        Result.Add "operator", "typeof"
        Result.Add "value", Value3
    ElseIf CurToken("type") = "symbol" Then
        If CurToken("value") = "{" Then
            Set Result = New Dictionary
            Result.Add "kind", "ArrayExpression"
            Result.Add "line", CurToken("line")
            Result.Add "start", CurToken("start")
            NextToken
            Dim Item As Dictionary
            Dim Items As New List
            Do
                Set Item = ExpectExpression
                If Item Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
                Items.Add Item
                If ExpectSymbol(",") Is Nothing Then Exit Do
            Loop
            If ExpectSymbol("}") Is Nothing Then ThrowError "Expected '}'.", CurToken("line"), CurToken("start")
'            Result.Add "type", Item("type")
            Result.Add "value", Items.Items
        ElseIf CurToken("value") = "+" Or CurToken("value") = "-" Then
            Set Result = New Dictionary
            Result.Add "kind", "UnaryExpression"
            Result.Add "line", CurToken("line")
            Result.Add "start", CurToken("start")
            Result.Add "operator", CurToken("value")
            Result.Add "prefix", True
            NextToken
            Dim Value2 As Dictionary
            Set Value2 = ExpectOneValue()
            If Value2 Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
            Result.Add "value", Value2
        ElseIf CurToken("value") = "++" Or CurToken("value") = "--" Then
            Set Result = New Dictionary
            Dim Operator As String
            Operator = CurToken("value")
            NextToken
            If CurToken("type") = "identifier" Then
                Set Result = New Dictionary
                Result.Add "kind", "UpdateExpression"
                Result.Add "line", CurToken("line")
                Result.Add "start", CurToken("start")
                Result.Add "operator", Operator
                Result.Add "prefix", True
                Result.Add "value", ExpectOneValue()
            Else
                ThrowError "Expected variable.", CurToken("line"), CurToken("start")
            End If
        ElseIf CurToken("value") = "!" Then
            Set Result = New Dictionary
            Result.Add "kind", "UnaryExpression"
            Result.Add "operator", CurToken("value")
            Result.Add "prefix", True
            NextToken
            Set Value = ExpectOneValue()
            If Value Is Nothing Then ThrowError "Expected value.", CurToken("line"), CurToken("start")
            Result.Add "value", Value
        Else
'            ThrowError "Unexpected token '" & CurToken("value") & "'.", CurToken("line"), CurToken("start")
        End If
'    ElseIf CurToken("type") = "keyword" And CurToken("value") = "new" Then
'        Set Result = New Dictionary
'        Result.Add "kind", "NewExpression"
'        NextToken
'        Dim Name As Dictionary
'        Set Name = ExpectIdentifier
'        Result.Add "value", CurToken("value")
'    Else
'        Set Result = ExpectFunctionCall
    End If
    If Result Is Nothing = False And CompareToken(CurToken, "symbol", "[") Then
        NextToken
        Set Result = ExpectArrayAccessExpression(Result)
    End If
    Set ExpectOneValue = Result
End Function

Private Function ExpectArrayAccessExpression(Expression As Dictionary) As Dictionary
    Dim Result As New Dictionary
    Dim Value As Dictionary
    Result.Add "kind", "ArrayAccessExpression"
    Result.Add "line", pTokens(CurTokenPos - 2)("line")
    Result.Add "start", pTokens(CurTokenPos - 2)("start")
    Set Value = ExpectExpression
    If Value Is Nothing Then ThrowError "Expected expression.", CurToken("line"), CurToken("start")
    If ExpectSymbol("]") Is Nothing Then ThrowError "Expetcted ']'.", CurToken("line"), CurToken("start")
    Result.Add "object", Expression
    Result.Add "index", Value
    If ExpectSymbol("[") Is Nothing = False Then Set Result = ExpectArrayAccessExpression(Result)
    Set ExpectArrayAccessExpression = Result
End Function

Private Function ExpectIdentifier() As Dictionary
    If CurToken("type") = "identifier" Then
        Set ExpectIdentifier = pTokens(CurTokenPos)
        NextToken
    End If
End Function

Private Function ExpectKeyword(Optional Value As String) As Dictionary
    If CurToken("type") = "keyword" And (Value = "" Or CurToken("value") = Value) Then
        Set ExpectKeyword = pTokens(CurTokenPos)
        NextToken
    End If
End Function

Private Function ExpectSymbol(Optional Value As String) As Dictionary
    If CurToken("type") = "symbol" And (Value = "" Or CurToken("value") = Value) Then
        Set ExpectSymbol = pTokens(CurTokenPos)
        NextToken
    End If
End Function

Private Function ExpectType(Optional IncludeVoid As Boolean = True) As String
    Dim PossibleType As Dictionary
    Set PossibleType = ExpectKeyword
    If Not PossibleType Is Nothing Then
        If InString("_any_short_int_long_float_double_bool_string_", "_" & PossibleType("value") & "_") Or (PossibleType("value") = "void" And IncludeVoid) Then
            ExpectType = PossibleType("value")
        Else
            PrevToken
        End If
    End If
End Function

Private Function ExpectModifier() As String
    If CurToken("type") = "keyword" And (CurToken("value") = "private" Or CurToken("value") = "public") Then
        ExpectModifier = pTokens(CurTokenPos)("value")
        NextToken
    End If
End Function

Public Function ParseLib(Tokens As List, LibClass As Object) As Dictionary
    Set pTokens = Tokens
    CurTokenPos = -1
    NextToken
    
    Dim Stmt As Dictionary
    Dim Funcs As New Dictionary
    Dim Consts As New Dictionary
    
    Do While CurTokenPos < pTokens.Length - 1
        Set Stmt = ExpectLibConstant
        If Stmt Is Nothing Then Set Stmt = ExpectLibFuncDef
        
        If Stmt("kind") = "Identifier" Then
            Consts.Add Stmt("name"), Stmt
        Else
            Stmt.Add "libclass", LibClass
            Funcs.Add Stmt("name"), Stmt
        End If
    Loop
    
    Set ParseLib = New Dictionary
    ParseLib.Add "functions", Funcs
    ParseLib.Add "constants", Consts
End Function

Private Function ExpectLibConstant() As Dictionary
    If ExpectKeyword("const") Is Nothing Then Exit Function
    
    Dim PossibleType As String
    Dim ArrayDim As Integer
    Dim PossibleName As Dictionary
    
    PossibleType = ExpectType(False)
    If PossibleType = "" Then ThrowError "Expected type.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
    
    ArrayDim = ExpectVariableArray
    
    Set PossibleName = ExpectIdentifier
    If PossibleName Is Nothing Then ThrowError "Expected constant name.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 2
    
    If ExpectSymbol("=") Is Nothing Then ThrowError "Constant must be initialized.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
    
    Dim Value As Dictionary
    Set Value = ValueConstant(ExpectOneValue)
    
    If Value Is Nothing Then ThrowError "Expected value.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
    
    ' todo: compare types
    
    Dim Constant As New Dictionary
    Constant.Add "kind", "Identifier"
    Constant.Add "name", PossibleName("value")
    Constant.Add "type", PossibleType
    Constant.Add "value", NewLiteral(PossibleType, Value("value"))
    Constant.Add "dim", ArrayDim
    
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';' at end of declaration.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Set ExpectLibConstant = Constant
End Function

Private Function ValueConstant(Value As Dictionary) As Dictionary
    If Value("kind") <> "Literal" Then
        If Value("kind") = "UnaryExpression" Then
            If Value("operator") = "+" Then
                Set ValueConstant = NewLiteral(Value("value")("type"), Value("value")("value"))
            ElseIf Value("operator") = "-" Then
                Set ValueConstant = NewLiteral(Value("value")("type"), -Value("value")("value"))
            Else
                ThrowError "Value must be constant.", CurToken("line"), CurToken("start")
            End If
        Else
            ThrowError "Value must be a constant.", CurToken("line"), CurToken("start")
        End If
    Else
        Set ValueConstant = Value
    End If
End Function

Private Function NewLiteral(ByVal LiteralType As String, LiteralValue As Variant) As Dictionary
    Set NewLiteral = New Dictionary
    NewLiteral.Add "kind", "Literal"
    NewLiteral.Add "type", LiteralType
    NewLiteral.Add "value", LiteralValue
End Function

Private Function ExpectLibFuncDef() As Dictionary
    Dim PossibleType As String
    Dim PossibleName As Dictionary
    Dim ArrayDim As Integer
    
    PossibleType = ExpectType
    ArrayDim = ExpectVariableArray
    Set PossibleName = ExpectIdentifier
    
    If PossibleType = "" Then ThrowError "Expected function type.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
    If PossibleName Is Nothing Then ThrowError "Expected function name.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
    If ExpectSymbol("(") Is Nothing Then ThrowError "Expected '('.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
    
    Dim Func As New Dictionary
    Dim Params As New List
    Func.Add "kind", "NativeFunction"
    Func.Add "name", PossibleName("value")
    Func.Add "type", mDeclTypes(PossibleType)
    Func.Add "dim", ArrayDim
    
    Dim Param As Dictionary
    Dim ReqParams As Long
    Dim HasOptionalParam As Boolean
    
    Do While ExpectSymbol(")") Is Nothing
        Dim Variable As Dictionary
        Set Variable = ExpectVariableDeclaration(True)
        If Variable Is Nothing Then ThrowError "Expected parameter type.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
        Variable("kind") = "Identifier"
        If Variable("init") Is Nothing Then
            ReqParams = ReqParams + 1
            If HasOptionalParam Then ThrowError "Optional parameters may only appear after required parameters", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
        Else
            HasOptionalParam = True
            Set Variable("init") = ValueConstant(Variable("init"))
            If Variable("init")("kind") <> "Literal" Then ThrowError "Default parameter value must be constant.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
            If Variable("dim") <> 0 Or Variable("rest") = True Then ThrowError "Default parameter cannot be array.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
            If GetTypeKind(Variable("init")("type")) <> GetTypeKind(Variable("type")) Then ThrowError "Type mismatch.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
        End If
        Params.Add Variable
        
        If ExpectSymbol(")") Is Nothing = False Then
            Exit Do
        End If
        
        If ExpectSymbol(",") Is Nothing Then
            ThrowError "Expected ',' or ')'.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
        ElseIf Variable("rest") Then
            ThrowError "Comma not permitted after rest parameter.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("start")
        End If
    Loop
    
    Func.Add "params", Params
    Func.Add "requiredParams", ReqParams
    
    If ExpectSymbol(";") Is Nothing Then ThrowError "Expected ';' at end of declaration.", pTokens(CurTokenPos - 1)("line"), pTokens(CurTokenPos - 1)("end") + 1
    
    Set ExpectLibFuncDef = Func
End Function

Private Function InString(String1 As String, String2 As String) As Boolean
    InString = InStr(1, String1, String2) > 0
End Function

Private Function GetTypeKind(ByVal ValueType As String) As String
    Select Case ValueType
    Case "byte", "short", "integer", "long"
        GetTypeKind = "integer"
    Case "float", "double"
        GetTypeKind = "float"
    Case Else
        GetTypeKind = ValueType
    End Select
End Function

Private Sub ThrowError(ErrMsg As String, Optional ByVal Line As Long = -1, Optional ByVal Col As Long = -1)
    Dim Msg As String
    Msg = "PARSING ERROR: " & ErrMsg & vbNewLine & "at line " & Line & " col " & Col & " at file '" & FileName & "'"
    If Line > 0 And Col > 0 Then Msg = Msg & vbNewLine & Lines(Line - 1) & vbNewLine & Space(Col - 1) & "^"
    RaiseEvent OnError(Msg)
    End
End Sub
