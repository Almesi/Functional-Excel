VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_IncludeFormatXLF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IIncludeFormat

'FileFormat:
'==FunctionHead and Body==:
' Gets compressed into 1 line:
' ReturnType Name(ArgType Arg1, ...)
'{
'    =LAMBDA(x, x*2)
'}


Private Reader As New std_IncludeFileReader
Private Seperator As String
Private WB As Workbook

Public Function Create(ByVal nSeperator As String, ByVal nWB As Workbook) As std_IncludeFormatXLF
    Set Create = New std_IncludeFormatXLF
    Call Create.Initialize(nSeperator, nWB)
End Function

Friend Sub Initialize(ByVal nSeperator As String, ByVal nWB As Workbook)
    Seperator = nSeperator
    Set WB = nWB
End Sub


Private Function IIncludeFormat_Include(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As VBIDE.VBComponent
    Set IIncludeFormat_Include = Nothing
    Call HandleFunctions(IIncludeFormat_Code(Source))
End Function

Private Function IIncludeFormat_AddQueue(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As Boolean
    IIncludeFormat_AddQueue = Includer.AddToQueue(Me, Source)
End Function

Private Function IIncludeFormat_CanHandle(ByVal Source As std_IncludeSource) As Boolean
    If Source.Value       = Empty    Then Exit Function
    If Source.SourceType <> "String" Then Exit Function

    Dim Format As String
    Format = Reader.FileFormat(Source.Value)
    IIncludeFormat_CanHandle = Format Like "xlf"
End Function

Private Function IIncludeFormat_Code(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Code = Reader.ReadFileCompilable(Source.Value)
End Function

Private Function IIncludeFormat_Name(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Name = Reader.GetFileName(Source.Value)
End Function


Private Sub HandleFunctions(ByVal Source As String)
    Dim Functions() As String
    Functions = SeperateFunctions(Source)
    Dim i As Long
    For i = 0 To Ubound(Functions)
        Call HandleFunction(Functions(i))
    Next i
End Sub

Private Function SeperateFunctions(ByVal Source As String) As String()
    Dim Lines()       As String : Lines = Split(Source, vbCrLf)
    Dim Size          As Long   : Size  = Ubound(Lines)
    Dim MergedSize    As Long   : MergedSize = -1
    Dim MergedLines() As String

    Dim Index As Long
    Do While Index =< Size
        Dim CurrentSize As Long
        CurrentSize = GetLastFunctionLine(Lines, Index)

        Dim i As Long
        Dim Text As String
        For i = Index To CurrentSize
            Text = Text & Lines(i)
        Next i
        MergedSize = MergedSize + 1
        ReDim Preserve MergedLines(MergedSize)
        MergedLines(MergedSize) = Text
        Text = Empty
        Index = CurrentSize + 1
    Loop
    SeperateFunctions = MergedLines
End Function

Private Sub HandleFunction(ByVal FuncString As String)
    Dim NameStart As Long   : NameStart = InStr(1, FuncString, " ") + 1
    Dim NameEnd   As Long   : NameEnd   = InStr(1, FuncString, "(") - 1
    Dim BodyStart As Long   : BodyStart = InStr(1, FuncString, "{") + 1
    Dim BodyEnd   As Long   : BodyEnd   = InStr(1, FuncString, "}") - 1
    Dim Name      As String : Name      = Mid(FuncString, NameStart, NameEnd - NameStart + 1)
    Dim Body      As String : Body      = Replace(Trim(Mid(FuncString, BodyStart, BodyEnd - BodyStart + 1)), ",", Seperator)

    Call AddNamedFunction(Name, Body)
End Sub

Private Sub AddNamedFunction(ByVal Name As String, ByVal Body As String)
    Call WB.Names.Add(Name, Body)
End Sub


Private Function GetLastFunctionLine(ByRef Arr() As String, ByVal Start As Long) As Long
    Dim HeadPattern As String : HeadPattern = "^\s*([A-Za-z_]\w*(?:\s+[A-Za-z_]\w*)*)\s+([A-Za-z_]\w*)\s*\([^)]*\)\s*$"
    Dim OpenPattern As String : OpenPattern = "^\s*\{\s*$"
    Dim ClosPattern As String : ClosPattern = "^\s*\}\s*$"
    Dim Size        As Long   : Size        = Ubound(Arr)

    If Start + 1 > Size Then
        Err.Raise vbObjectError + 1000, _
            "std_IncludeFormatXLF.GetLastFunctionLine", _
            "Unexpected end of file after function header."
    End If


    If Match(HeadPattern, Arr(Start)) And Match(OpenPattern, Arr(Start + 1)) Then
        Dim i As Long
        For i = Start + 2 To Size
            If Match(ClosPattern, Arr(i)) Then
                GetLastFunctionLine = i
                Exit Function
            End If
        Next i
    End If

    Err.Raise vbObjectError + 1000, "std_IncludeFormatXLF.GetLastFunctionLine", _
          "Could not find a valid XLF function starting at line " & Start
End Function

Private Function Match(ByVal Pattern As String, ByVal Inputt As String) As Boolean
    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")

    With Regex
        .Pattern = Pattern
        .Global = False
        .IgnoreCase = True
    End With
    Match = Regex.Test(Inputt)
End Function