Attribute VB_Name = "Include"

Option Explicit

Public Sub IncludeFunctions()
    Dim Destination As IDestination         : Set Destination = std_ImmiedeateDestination.Create()
    Dim Handler     As std_ErrorHandler     : Set Handler     = std_ErrorHandler.Create(Destination, Destination)
    Dim Includer    As std_Include          : Set Includer    = std_Include.Create(ThisWorkbook.VBProject, Handler)
    Dim Format      As std_IncludeFormatXLF : Set Format      = std_IncludeFormatXLF.Create(",", ThisWorkbook)
    Dim Path        As String               : Let Path        = "C:\Users\deallulic\Documents\Functional Excel\Src"
    Includer.Recursive = True
    Includer.Log       = True
    Includer.Overwrite = True
    Includer.Increment = False
    Includer.Temporary = False

    Call Includer.AddFormat(Format)

    If Includer.IncludeFolder(Path) Then
        Debug.Print "Success: ", Includer.Build()
    End If
End Sub