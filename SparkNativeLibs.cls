VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SparkNativeLibs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Spark As SparkCore

Public Function GetDefinitions(LibId As String)
    Dim Defs As String
    
    Select Case LibId
    Case "io"
        Defs = Defs & "const short FATTR_NORMAL = 0;"
        Defs = Defs & "const short FATTR_READ_ONLY = 1;"
        Defs = Defs & "const short FATTR_HIDDEN = 2;"
        Defs = Defs & "const short FATTR_SYSTEM = 4;"
        Defs = Defs & "const short FATTR_VOLUME = 8;"
        Defs = Defs & "const short FATTR_DIRECTORY = 16;"
        Defs = Defs & "const short FATTR_ALIAS = 64;"
        Defs = Defs & "void print(any... text);"
        Defs = Defs & "void println(any... text);"
        Defs = Defs & "string readTextFile(string path, string charset = ""utf-8"");"
        Defs = Defs & "void writeTextFile(string path, string content, string charset = ""utf-8"");"
        Defs = Defs & "void deleteFile(string path);"
        Defs = Defs & "void renameFile(string path, string newPath);"
        Defs = Defs & "bool fileExists(string path);"
        Defs = Defs & "string dir(string pathName = """", short attributes = 0);"
    Case "math"
        Defs = Defs & "double sin(double value);"
        Defs = Defs & "double cos(double value);"
        Defs = Defs & "double tan(double value);"
        Defs = Defs & "double atn(double value);"
        Defs = Defs & "double abs(double value);"
        Defs = Defs & "int floor(double value);"
        Defs = Defs & "int ceil(double value);"
        Defs = Defs & "int round(double value);"
        Defs = Defs & "int sign(double value);"
        Defs = Defs & "double sqrt(double value);"
        Defs = Defs & "int parseInt(string value);"
        Defs = Defs & "double parseFloat(string value);"
        Defs = Defs & "double min(double a, double b);"
        Defs = Defs & "double max(double a, double b);"
        Defs = Defs & "double random();"
    Case "string"
        Defs = Defs & "string toString(any value);"
        Defs = Defs & "int strLen(string s);"
        Defs = Defs & "string[] split(string s, string delimiter);"
        Defs = Defs & "string join(string[] arr, string delimiter = """");"
        Defs = Defs & "short getCharCode(string char);"
        Defs = Defs & "string getCharFromCode(short charCode);"
        Defs = Defs & "string substr(string s, int start, int length = -1);"
        Defs = Defs & "string substring(string s, int start, int end = -1);"
        Defs = Defs & "bool startsWith(string s, string searchString, int pos = 0);"
        Defs = Defs & "bool endsWith(string s, string searchString, int endPos = -1);"
    Case "sys"
        Defs = Defs & "int msgbox(string prompt, int buttons, string title);"
        Defs = Defs & "float timer();"
        Defs = Defs & "int doEvents();"
        Defs = Defs & "bool isKeyPressed(int keyCode);"
    Case "ppt"
        Defs = Defs & "const short SHAPE_RECTANGLE = 1;"
        Defs = Defs & "const short SHAPE_ISOSCELES_TRIANGLE = 7;"
        Defs = Defs & "const short SHAPE_OVAL = 9;"
        Defs = Defs & "void addShape(int slideIndex, short shapeType, string shapeName, float left, float top, float width, float height);"
        Defs = Defs & "void setShapeLeft(int slideIndex, string shapeName, float left);"
        Defs = Defs & "void setShapeTop(int slideIndex, string shapeName, float top);"
        Defs = Defs & "void setShapeWidth(int slideIndex, string shapeName, float width);"
        Defs = Defs & "void setShapeHeight(int slideIndex, string shapeName, float height);"
        Defs = Defs & "void setShapeRotation(int slideIndex, string shapeName, float rotation);"
        Defs = Defs & "void setShapeText(int slideIndex, string shapeName, string text);"
        Defs = Defs & "void setShapeFillColor(int slideIndex, string shapeName, int color);"
        Defs = Defs & "int getShapeFillColor(int slideIndex, string shapeName);"
        Defs = Defs & "void setShapeFillTransparency(int slideIndex, string shapeName, float transparency);"
        Defs = Defs & "float getShapeFillTransparency(int slideIndex, string shapeName);"
        Defs = Defs & "void setShapeFillPicture(int slideIndex, string shapeName, string pictureFile);"
        Defs = Defs & "void setShapeLineColor(int slideIndex, string shapeName, int color);"
        Defs = Defs & "int getShapeLineColor(int slideIndex, string shapeName);"
        Defs = Defs & "void setShapeLineTransparency(int slideIndex, string shapeName, float transparency);"
        Defs = Defs & "float getShapeLineTransparency(int slideIndex, string shapeName);"
        Defs = Defs & "float getShapeLeft(int slideIndex, string shapeName);"
        Defs = Defs & "float getShapeTop(int slideIndex, string shapeName);"
        Defs = Defs & "float getShapeWidth(int slideIndex, string shapeName);"
        Defs = Defs & "float getShapeHeight(int slideIndex, string shapeName);"
        Defs = Defs & "float getShapeRotation(int slideIndex, string shapeName);"
        Defs = Defs & "void setShapeAction(int slideIndex, string shapeName, short mouseActivation, short action);"
        Defs = Defs & "void setShapeActionHyperlinkAddress(int slideIndex, string shapeName, short mouseActivation, string hyperlinkAddress);"
        Defs = Defs & "void setShapeActionRunMacro(int slideIndex, string shapeName, short mouseActivation, string macroName);"
        Defs = Defs & "string getShapeText(int slideIndex, string shapeName);"
        Defs = Defs & "string getShapeName(int slideIndex, int shapeIndex);"
        Defs = Defs & "void deleteShape(int slideIndex, string shapeName);"
        Defs = Defs & "void duplicateShape(int slideIndex, string shapeName);"
        Defs = Defs & "int getShapeCount(int slideIndex);"
        Defs = Defs & "int getSlideCount();"
        Defs = Defs & "int getCurrentSlide();"
        Defs = Defs & "string[] getShapes(int slideIndex);"
        Defs = Defs & "void gotoSlide(int slideIndex);"
    End Select
    
    GetDefinitions = Defs
End Function

' ###############
' #  LIBRARIES  #
' ###############


' ============
' io
' ============

' void print(any... text)
Public Sub sparklib_print(Text As Variant)
'    Spark.Throw "An error has occured."
    If IsEmpty(Text(0)) Then Exit Sub
    Dim i As Long
    For i = 0 To UBound(Text)
        Spark.LogMsg Spark.EvalToString(Text(i)) & IIf(i < UBound(Text), vbTab, "")
    Next
End Sub

' void println(any... text)
Public Sub sparklib_println(Text As Variant)
    If IsEmpty(Text(0)) Then Exit Sub
    sparklib_print Text
    Spark.LogMsg vbNewLine
End Sub

' string readTextFile(string path, string charset = "utf-8")
Public Function sparklib_readTextFile(Path As String, Charset As String) As String
    sparklib_readTextFile = ReadTextFile(Path, Charset)
End Function

' void writeTextFile(string path, string content, string charset = "utf-8")
Public Sub sparklib_writeTextFile(Path As String, Content As String, Charset As String)
    WriteTextFile Path, Content, Charset
End Sub

' void deleteFile(string path)
Public Sub sparklib_deleteFile(Path As String)
    Kill Path
End Sub

' void renameFile(string path, string newPath)
Public Sub sparklib_renameFile(Path As String, NewPath As String)
    Name Path As NewPath
End Sub

' bool fileExists(string path)
Public Function sparklib_fileExists(Path As String) As Boolean
    sparklib_fileExists = Dir(Path) <> ""
End Function

' string dir(string pathName, short attributes)
Public Function sparklib_dir(PathName As String, Attributes As Integer) As String
    If PathName = "" Then
        sparklib_dir = Dir(, Attributes)
    Else
        sparklib_dir = Dir(PathName, Attributes)
    End If
End Function


' ============
' math
' ============

' double sin(double value)
Public Function sparklib_sin(Value As Double) As Double
    sparklib_sin = Sin(Value)
End Function

' double cos(double value)
Public Function sparklib_cos(Value As Double) As Double
    sparklib_cos = Cos(Value)
End Function

' double tan(double value)
Public Function sparklib_tan(Value As Double) As Double
    sparklib_tan = Tan(Value)
End Function

' double atn(double value)
Public Function sparklib_atn(Value As Double) As Double
    sparklib_atn = Atn(Value)
End Function

' double abs(double value)
Public Function sparklib_abs(Value As Double) As Double
    sparklib_abs = Abs(Value)
End Function

' int floor(double value)
Public Function sparklib_floor(Value As Double) As Long
    sparklib_floor = Int(Value)
End Function

' int ceil(double value)
Public Function sparklib_ceil(Value As Double) As Long
    Dim v As Long
    v = Int(Value)
    If v = Value Then
        sparklib_ceil = v
    Else
        sparklib_ceil = v + 1
    End If
End Function

' int round(double value)
Public Function sparklib_round(Value As Double) As Long
    sparklib_round = Round(Value)
End Function

' int sign(double value)
Public Function sparklib_sign(Value As Double) As Long
    sparklib_sign = Sgn(Value)
End Function

' double sqrt(double value)
Public Function sparklib_sqrt(Value As Double) As Double
    sparklib_sqrt = Sqr(Value)
End Function

' int parseInt(string value)
Public Function sparklib_parseInt(Value As String) As Long
    sparklib_parseInt = CInt(Replace(Value, ".", Format(0, ".")))
End Function

' double parseFloat(string value)
Public Function sparklib_parseFloat(Value As String) As Double
    sparklib_parseFloat = CDbl(Replace(Value, ".", Format(0, ".")))
End Function

' double min(double a, double b)
Public Function sparklib_min(a As Double, B As Double) As Double
    sparklib_min = IIf(a < B, a, B)
End Function

' double max(double a, double b)
Public Function sparklib_max(a As Double, B As Double) As Double
    sparklib_max = IIf(a > B, a, B)
End Function

' double random()
Public Function sparklib_random() As Double
    Randomize
    sparklib_random = Rnd
End Function


' ============
' string
' ============

' string toString(any value)
Public Function sparklib_toString(Value As Variant) As String
    sparklib_toString = Spark.EvalToString(Value)
End Function

' int strLen(string s)
Public Function sparklib_strLen(s As String) As Long
    sparklib_strLen = Len(s)
End Function

' string[] split(string s, string delimiter)
Public Function sparklib_split(s As String, Delimiter As String) As String()
    sparklib_split = Split(s, Delimiter)
End Function

' string join(string[] arr, string delimiter = "")
Public Function sparklib_join(Arr() As Variant, Delimiter As String) As String
    sparklib_join = Join(Arr, Delimiter)
End Function

' short getCharCode(string char)
Public Function sparklib_getCharCode(Char As String) As Integer
    sparklib_getCharCode = AscW(Char)
End Function

' string getCharFromCode(short charCode)
Public Function sparklib_getCharFromCode(Char As Integer) As String
    sparklib_getCharFromCode = ChrW(Char)
End Function

' string substr(string s, int start, int length = -1)
Public Function sparklib_substr(s As String, Start As Long, Length As Long) As String
    Start = Start + 1
    If Length = -1 Then Length = Len(s) - Start + 1
    sparklib_substr = Mid(s, Start, Length)
End Function

' string substring(string s, int start, int end)
Public Function sparklib_substring(s As String, Start As Long, EndPos As Long) As String
    Start = Start + 1
    If EndPos = -1 Then
        EndPos = Len(s)
    Else
        EndPos = EndPos + 1
    End If
    sparklib_substring = Mid(s, Start, EndPos - Start + 1)
End Function

' bool startsWith(string s, string searchString, int pos)
Public Function sparklib_startsWith(s As String, SearchString As String, Pos As Long) As Boolean
    sparklib_startsWith = Mid(s, Pos + 1, Len(SearchString)) = SearchString
End Function

' bool startsWith(string s, string searchString, int endPos)
Public Function sparklib_endsWith(s As String, SearchString As String, EndPos As Long) As Boolean
    If EndPos = -1 Then EndPos = Len(s) - 1
    sparklib_endsWith = Mid(s, EndPos - Len(SearchString) + 2, Len(SearchString)) = SearchString
End Function


' ============
' sys
' ============

' int msgbox(string prompt, int buttons, int title)
Public Function sparklib_msgBox(Prompt As String, Buttons As Long, Title As String) As Long
    Dim Res As VbMsgBoxResult
    Res = MsgBox(Prompt, Buttons, Title)
    sparklib_msgBox = Res
End Function

' float timer()
Public Function sparklib_timer() As Single
    sparklib_timer = Timer
End Function

' double getSysTime()
Public Function sparklib_getSysTime() As Double
    sparklib_getSysTime = GetTickCount / 1000
End Function

' int doEvents()
Public Function sparklib_doEvents() As Integer
    sparklib_doEvents = DoEvents
End Function

' bool isKeyPressed(int keyCode)
Public Function sparklib_isKeyPressed(keyCode As Long) As Boolean
    sparklib_isKeyPressed = GetAsyncKeyState(keyCode)
End Function

' short getDate()
Public Function sparklib_getDate() As Integer
    sparklib_getDate = Day(Date)
End Function

' short getMonth()
Public Function sparklib_getMonth() As Integer
    sparklib_getMonth = Month(Date) - 1
End Function

' short getYear()
Public Function sparklib_getYear() As Integer
    sparklib_getYear = Year(Date)
End Function

' short getHours()
Public Function sparklib_getHours() As Integer
    sparklib_getHours = Hour(Time)
End Function

' short getMinutes()
Public Function sparklib_getMinutes() As Integer
    sparklib_getMinutes = Minute(Time)
End Function

' short getSeconds()
Public Function sparklib_getSeconds() As Integer
    sparklib_getSeconds = Second(Time)
End Function

' short getMillis()
Public Function sparklib_getMillis() As Integer
    Dim SysTime As SYSTEMTIME
    GetSystemTime SysTime
    sparklib_getMillis = SysTime.wMilliseconds
End Function

' long getTime()
Public Function sparklib_getTime() As LongLong
    Dim SysTime As SYSTEMTIME
    GetSystemTime SysTime
    sparklib_getTime = (Date - #1/1/1970#) * 86400000 + SysTime.wMilliseconds
End Function


' ============
' ppt
' ============

' void addShape(int slideIndex, short shapeType, string shapeName, float left, float top, float width, float height)
Public Sub sparklib_addShape(SlideIndex As Long, ShapeType As Integer, ShapeName As String, Left As Single, Top As Single, Width As Single, Height As Single)
    ActivePresentation.Slides(SlideIndex).Shapes.AddShape(ShapeType, Left, Top, Width, Height).Name = ShapeName
End Sub

' void setShapeLeft(int slideIndex, string shapeName, float left)
Public Sub sparklib_setShapeLeft(SlideIndex As Long, ShapeName As String, Left As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Left = Left
End Sub

' void setShapeTop(int slideIndex, string shapeName, float top)
Public Sub sparklib_setShapeTop(SlideIndex As Long, ShapeName As String, Top As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Top = Top
End Sub

' void setShapeWidth(int slideIndex, string shapeName, float width)
Public Sub sparklib_setShapeWidth(SlideIndex As Long, ShapeName As String, Width As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Width = Width
End Sub

' void setShapeHeight(int slideIndex, string shapeName, float height)
Public Sub sparklib_setShapeHeight(SlideIndex As Long, ShapeName As String, Height As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Height = Height
End Sub

' void setShapeRotation(int slideIndex, string shapeName, float rotation)
Public Sub sparklib_setShapeRotation(SlideIndex As Long, ShapeName As String, Rotation As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Rotation = Rotation
End Sub

' void setShapeText(int slideIndex, string shapeName, string text)
Public Sub sparklib_setShapeText(SlideIndex As Long, ShapeName As String, Text As String)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).TextFrame.TextRange.Text = Text
End Sub

' void setShapeFillColor(int slideIndex, string shapeName, int color)
Public Sub sparklib_setShapeFillColor(SlideIndex As Long, ShapeName As String, Color As Long)
    With ActivePresentation.Slides(SlideIndex).Shapes(ShapeName)
        .Fill.Solid
        .Fill.ForeColor.RGB = Color
    End With
End Sub

' int getShapeFillColor(int slideIndex, string shapeName)
Public Function sparklib_getShapeFillColor(SlideIndex As Long, ShapeName As String) As Long
    sparklib_getShapeFillColor = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Fill.ForeColor.RGB = Color
End Function

' void setShapeFillTransparency(int slideIndex, string shapeName, float transparency)
Public Sub sparklib_setShapeFillTransparency(SlideIndex As Long, ShapeName As String, Transparency As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Fill.Transparency = Transparency
End Sub

' float getShapeFillTransparency(int slideIndex, string shapeName)
Public Function sparklib_getShapeFillTransparency(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeFillTransparency = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Fill.Transparency
End Function

' void setShapeFillPicture(int slideIndex, string shapeName, string pictureFile)
Public Sub sparklib_setShapeFillPicture(SlideIndex As Long, ShapeName As String, PictureFile As String)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Fill.UserPicture PictureFile
End Sub

' void setShapeLineColor(int slideIndex, string shapeName, int color)
Public Sub sparklib_setShapeLineColor(SlideIndex As Long, ShapeName As String, Color As Long)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Line.ForeColor.RGB = Color
End Sub

' int getShapeLineColor(int slideIndex, string shapeName)
Public Function sparklib_getShapeLineColor(SlideIndex As Long, ShapeName As String) As Long
    sparklib_getShapeLineColor = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Line.ForeColor.RGB = Color
End Function

' void setShapeLineTransparency(int slideIndex, string shapeName, float transparency)
Public Sub sparklib_setShapeLineTransparency(SlideIndex As Long, ShapeName As String, Transparency As Single)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Line.Transparency = Transparency
End Sub

' float getShapeLineTransparency(int slideIndex, string shapeName)
Public Function sparklib_getShapeLineTransparency(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeLineTransparency = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Line.Transparency
End Function

' float getShapeLeft(int slideIndex, string shapeName)
Public Function sparklib_getShapeLeft(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeLeft = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Left
End Function

' float getShapeTop(int slideIndex, string shapeName)
Public Function sparklib_getShapeTop(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeTop = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Top
End Function

' float getShapeWidth(int slideIndex, string shapeName)
Public Function sparklib_getShapeWidth(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeWidth = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Width
End Function

' float getShapeHeight(int slideIndex, string shapeName)
Public Function sparklib_getShapeHeight(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeHeight = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Height
End Function

' float getShapeRotation(int slideIndex, string shapeName)
Public Function sparklib_getShapeRotation(SlideIndex As Long, ShapeName As String) As Single
    sparklib_getShapeRotation = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Rotation
End Function

' string getShapeText(int slideIndex, string shapeName)
Public Function sparklib_getShapeText(SlideIndex As Long, ShapeName As String) As String
    sparklib_getShapeText = ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).TextFrame.TextRange.Text
End Function

' void setShapeAction(int slideIndex, string shapeName, short mouseActivation, short action)
Public Sub sparklib_setShapeAction(SlideIndex As Long, ShapeName As String, MouseActivation As Integer, Action As Integer)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).ActionSettings(MouseActivation).Action = Action
End Sub

' void setShapeActionHyperlinkAddress(int slideIndex, string shapeName, short mouseActivation, string hyperlinkAddress)
Public Sub sparklib_setShapeActionHyperlinkAddress(SlideIndex As Long, ShapeName As String, MouseActivation As Integer, HyperlinkAddress As String)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).ActionSettings(MouseActivation).Hyperlink.Address = HyperlinkAddress
End Sub

' void setShapeActionRunMacro(int slideIndex, string shapeName, short mouseActivation, string macroName)
Public Sub sparklib_setShapeActionRunMacro(SlideIndex As Long, ShapeName As String, MouseActivation As Integer, MacroName As String)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).ActionSettings(MouseActivation).Run = MacroName
End Sub

' string getShapeName(int slideIndex, int shapeIndex)
Public Function sparklib_getShapeName(SlideIndex As Long, ShapeIndex As Long) As String
    sparklib_getShapeName = ActivePresentation.Slides(SlideIndex).Shapes(ShapeIndex).Name
End Function

' void deleteShape(int slideIndex, string shapeName)
Public Sub sparklib_deleteShape(SlideIndex As Long, ShapeName As String)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Delete
End Sub

' void duplicateShape(int slideIndex, string shapeName)
Public Sub sparklib_duplicateShape(SlideIndex As Long, ShapeName As String)
    ActivePresentation.Slides(SlideIndex).Shapes(ShapeName).Duplicate
End Sub

' int getShapeCount(int slideIndex)
Public Function sparklib_getShapeCount(SlideIndex As Long) As Long
    sparklib_getShapeCount = ActivePresentation.Slides(SlideIndex).Shapes.Count
End Function

' int getSlideCount()
Public Function sparklib_getSlideCount() As Long
    sparklib_getSlideCount = ActivePresentation.Slides.Count
End Function

' int getCurrentSlide()
Public Function sparklib_getCurrentSlide() As Long
    sparklib_getCurrentSlide = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
End Function

' string[] getShapes(int slideIndex)
Public Function sparklib_getShapes(SlideIndex As Long) As String()
    Dim Shp As Shape
    Dim Shps As ShapeRange
    Dim Res() As String
    Dim i As Long
    Set Shps = ActivePresentation.Slides(SlideIndex).Shapes.Range
    If Shps.Count > 0 Then
        ReDim Res(Shps.Count - 1) As String
        For i = 0 To Shps.Count - 1
            Res(i) = Shps(i + 1).Name
        Next
    End If
    sparklib_getShapes = Res
End Function

' void gotoSlide(int slideIndex)
Public Sub sparklib_gotoSlide(SlideIndex As Long)
    ActivePresentation.SlideShowWindow.View.GotoSlide SlideIndex
End Sub
