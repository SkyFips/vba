Attribute VB_Name = "mImportModules"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: within "listOfModules" all modules (without ending)              |
'              can be named and will be used to import into an excel sheet      |
'                                                                               |
' ##############################################################################|
'   version | description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/
Dim oFileSys  As Object
Dim oFolder   As Object
Dim oFile     As Object
Dim found     As Boolean
Dim oModules  As Collection
Dim done      As Collection

Private Property Get fileSys() As Object
  If oFileSys Is Nothing Then Set oFileSys = CreateObject("Scripting.FileSystemObject")
  Set fileSys = oFileSys
End Property

Private Sub ImportModules()
  On Error GoTo Error
  Dim c As Object
  Dim i As Integer
  Dim x As Byte
  x = 10
  Set oModules = Nothing
  Do While modules.count > 0 And x > 0
    Set done = New Collection
    For i = 1 To modules.count
      If moduleExist(modules.Item(i)) Then
        Err.Raise 1, _
                  "mImportModules.ImportModules", _
                  "execute ""RemoveModule"" before import"
      End If
      Debug.Print "[INFO] import """ & modules.Item(i) & """"
      Call SearchModule(searchFolder, modules.Item(i))
      If found Then
        Debug.Print "[INFO]  import " & oFile.path
        ThisWorkbook.VBProject.VBComponents.Import oFile.path
        found = False
        done.Add i
      End If
    Next i
    x = x - 1
    For i = done.count To 1 Step -1
      Call modules.Remove(done(i))
    Next i
    Dim n As String
    n = vbNullString
    Dim s As Variant
    For Each s In modules
      n = n & " - " & s & vbLf
    Next s
    If modules.count > 0 Then
      If Not MsgBox("following components haven't been imported (no file found)" & vbLf & _
                    n, _
                    vbRetryCancel, _
                    "not imported components") = vbRetry Then Exit Do
    End If
    Set searchFolder = Nothing
  Loop
  Exit Sub
Error:
  MsgBox "error:  " & Err.Number & vbLf & _
         "source: " & Err.Source & vbLf & _
         "descr.:   " & Err.description, _
         vbOKOnly, _
         "import modules"
End Sub

Private Sub RemoveModules()
  Dim v As Variant
  Dim c As Object
  Set oModules = Nothing
  For Each v In modules
    For Each c In ThisWorkbook.VBProject.VBComponents
      If c.name = v And Not v = "mImportModules" Then
        Debug.Print "[INFO]  remove """ & c.name & """"
        ThisWorkbook.VBProject.VBComponents.Remove c
        Exit For
      End If
    Next c
  Next v
End Sub

Private Function moduleExist(n As String) As Boolean
  Dim c As Object
  For Each c In ThisWorkbook.VBProject.VBComponents
    If c.name = n Then
      moduleExist = True
      Exit For
    End If
  Next c
End Function

Private Function extension(c As Object) As String
  ' input as "VBIDE.VBComponent"
  Select Case c.Type 'vbext_ComponentType
  Case 1 'vbext_ct_StdModule
    extension = "bas"
  Case 2 'vbext_ct_ClassModule
    extension = "cls"
  Case 3 'vbext_ct_MSForm
    extension = "frm"
  End Select
End Function

Private property get modules() As Collection
  If oModules Is Nothing Then
    Set oModules = New Collection
    Dim v As Variant
    For Each v In listOfModules
      oModules.Add v
    Next v
  End If
  Set modules = oModules
End Function

Private Sub SearchModule(f As Object, n As String)
  Dim subFolder As Object
  Dim file      As Object
  Dim x         As String
  For Each file In f.Files
    If file.Name = n Or Split(file.Name, ".")(0) = n Then
      x = Split(file.Name, ".")(UBound(Split(file.Name, ".")))
      If x = "bas" Or x = "cls" Or x = "frm" Then
        Set oFile = file
        found = True
        Exit For
      End If
    End If
  Next
  If Not found Then
    For Each subFolder In f.SubFolders
      Call SearchModule(subFolder, n)
      If found Then Exit For
    Next
  End If
End Sub

Private Property Get searchFolder() As Object
  If oFolder Is Nothing Then
    With Application.FileDialog(msoFileDialogFolderPicker)
      .title = "folder for modules"
      If .Show Then
        Dim p As String
        Set oFolder = fileSys.GetFolder(.SelectedItems(1))
      Else
        Err.Raise 79, _
                  "mExportImportModules.searchFolder", _
                  "no folder selected"
      End If
    End With
  End If
  Set searchFolder = oFolder
End Property

Private Property Set searchFolder(f As Object)
  Set oFolder = f
End Property

' modules to import
Private Property Get listOfModules() As Collection
  Set listOfModules = new Collection
  ' listOfModules.Add "cFooBar"
End Property
