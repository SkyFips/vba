VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cComparer_semanticVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' (c)2018 Scientific Games International GmbH                                   |
'                                                                               |
' Author(s):      Philipp Gorkiewicz                                            |
' Version:        2.1.0                                                         |
' Creation date:  2018-09-23                                                    |
' Last change:    2022-04-02                                                    |
'                                                                               |
' Description:    implementation for IComparer for Versions                     |
'                 a version must look like the following                        |
'                 a.b.c-prereleas+metadata (https://semver.org)                 |
'                 a = Major Version                                             |
'                 b = Minor Version                                             |
'                 c = Patch Version                                             |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
'       1.0.1 | implement semantic versioning                                   |
'       1.0.2 | add logging messages                                            |
'       1.0.3 | renaming of variables                                           |
'       1.1.0 | remove globLogger                                               |
'       2.0.0 | make verify public and do not fail anymore in case of invalid   |
'               version input, rename class                                     |
'       2.1.0 | use local logger                                                |
'       2.1.1 | remove unused variable                                          |
' ##############################################################################/
Implements iComparer

'comparison pattern from semver.org
Const COMPARISON_PATTERN  As String = "^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-((?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*)(?:\.(?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*))*))?(?:\+([0-9a-zA-Z-]+(?:\.[0-9a-zA-Z-]+)*))?$"

Dim oLogger As iLogger
Dim oRegEx  As New RegExp

Private Enum version
  major = 0
  minor = 1
  patch = 2
  preRelease = 3
  metaData = 4
End Enum
Private Sub Class_Initialize()
  'referenceCheck.CheckReference "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5 'Microsoft VBScript Regular Expressions 5.5
End Sub
Private Property Get logger() As iLogger
  If oLogger Is Nothing Then Set oLogger = New cLogger_watchWindow
  Set logger = oLogger
End Property

Public Function isPreRelease(version As String) As Boolean
  If Not IsEmpty(regEx(version).SubMatches(preRelease)) Then isPreRelease = True
  If iComparer_compare(version, "1.0.0") = less Then isPreRelease = True
End Function

Public Function hasMetaData(version As String) As Boolean
  If Not IsEmpty(regEx(version).SubMatches(metaData)) Then hasMetaData = True
End Function
Private Function regEx(version As String) As Match
  version = Trim(CStr(version))
  Call verify(version)
  oRegEx.Pattern = COMPARISON_PATTERN
  Set regEx = oRegEx.Execute(version).Item(0)
End Function
Private Function iComparer_compare(input_A As Variant, input_B As Variant) As compareResult
  Dim versionA  As String
  Dim versionB  As String
  Dim matchA    As Match
  Dim matchB    As Match
  Dim result    As compareResult
  Dim regEx     As New RegExp
    
  versionA = Trim(CStr(input_A))
  versionB = Trim(CStr(input_B))
  If Not verify(versionA) Then
    iComparer_compare = less
    Exit Function
  ElseIf Not verify(versionB) Then
    iComparer_compare = greater
    Exit Function
  End If
  
  regEx.Pattern = COMPARISON_PATTERN
  Set matchA = regEx.Execute(versionA).Item(0)
  Set matchB = regEx.Execute(versionB).Item(0)

  result = compare(matchA.SubMatches(version.major), matchB.SubMatches(version.major))
  If result = equal Then
    result = compare(matchA.SubMatches(version.minor), matchB.SubMatches(version.minor))
    If result = equal Then
      result = compare(matchA.SubMatches(version.patch), matchB.SubMatches(version.patch))
      If result = equal Then
        result = compare(matchA.SubMatches(version.preRelease), matchB.SubMatches(version.preRelease))
      End If
    End If
  End If
  iComparer_compare = result
End Function

Private Function verify(versionInput As String) As Boolean
  Dim regEx As New RegExp
  
  regEx.Pattern = COMPARISON_PATTERN
  verify = regEx.Test(versionInput)
  If Not verify Then logger.log FAILURE, """" & versionInput & """ is not a valid semantic version"
End Function

Private Function compare(stringA As String, stringB As String) As compareResult
  If IsNumeric(stringA) And IsNumeric(stringB) Then
    If CLng(stringA) > CLng(stringB) Then
      compare = greater
    ElseIf CLng(stringA) < CLng(stringB) Then
      compare = less
    Else
      compare = equal
    End If
  Else
    If stringA = vbNullString And Not (stringB = vbNullString) Then
      compare = greater
    ElseIf Not (stringA = vbNullString) And stringB = vbNullString Then
      compare = less
    Else
      If stringA > stringB Then
        compare = greater
      ElseIf stringA < stringB Then
        compare = less
      Else
        compare = equal
      End If
    End If
  End If
End Function
