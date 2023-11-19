VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cComparer_colorHLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As LongPtr, ByRef pwHue As LongPtr, ByRef pwLuminance As LongPtr, ByRef pwSaturation As LongPtr) As LongPtr
Private Declare PtrSafe Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As LongPtr
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: implementation for IComparer for HLS comparison                  |
'                                                                               |
' ##############################################################################|
' Version | Description                                                         |
'   1.0.0 | Initial Version                                                     |
' ##############################################################################/
Implements iComparer

Private Type typRGB
  red   As LongPtr
  green As LongPtr
  blue  As LongPtr
End Type
Private Type typHLS
  hue         As LongPtr
  saturation  As LongPtr
  luminance   As LongPtr
End Type

Private Function iComparer_compare(input_A As Variant, _
                                   input_B As Variant) As compareResult
  Dim a As typHLS
  Dim b As typHLS
  ColorRGBToHLS CLng(input_A), a.hue, a.luminance, a.saturation
  ColorRGBToHLS CLng(input_B), b.hue, b.luminance, b.saturation
  iComparer_compare = compare(a, b)
End Function

Private Function convertLongToRGB(lngInput As LongPtr) As typRGB
  convertLongToRGB.red = lngInput Mod 256
  convertLongToRGB.green = (lngInput \ 256) Mod 256
  convertLongToRGB.blue = (lngInput \ 256 \ 256) Mod 256
End Function
Private Function compare(a As typHLS, b As typHLS) As compareResult
  If a.hue < b.hue Then
    compare = greater
  ElseIf a.hue > b.hue Then
    compare = less
  Else
    If a.saturation > b.saturation Then
      compare = greater
    ElseIf a.saturation < b.saturation Then
      compare = less
    Else
      If a.luminance < b.luminance Then
        compare = greater
      ElseIf a.luminance > b.luminance Then
        compare = less
      Else
        compare = equal
      End If
    End If
  End If
End Function
