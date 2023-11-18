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
' (c)2020 Scientific Games International GmbH                                   |
'                                                                               |
' Author(s):      Philipp Gorkiewicz                                            |
' Version:        1.0.1                                                         |
' Creation date:  2020-04-08                                                    |
' Last change:    2022-09-20                                                    |
'                                                                               |
' Description:    implementation for IComparer for HLS comparison               |
'                                                                               |
' ##############################################################################|
'     Version | Description                                                     |
'       1.0.0 | Initial Version                                                 |
'       1.0.1 | remove unused variable MODULE                                   |
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
  Dim colorA  As typHLS
  Dim colorB  As typHLS
  ColorRGBToHLS CLng(input_A), colorA.hue, colorA.luminance, colorA.saturation
  ColorRGBToHLS CLng(input_B), colorB.hue, colorB.luminance, colorB.saturation
  iComparer_compare = compare(colorA, colorB)
End Function

Private Function convertLongToRGB(lngInput As LongPtr) As typRGB
  convertLongToRGB.red = lngInput Mod 256
  convertLongToRGB.green = (lngInput \ 256) Mod 256
  convertLongToRGB.blue = (lngInput \ 256 \ 256) Mod 256
End Function
Private Function compare(colorA As typHLS, colorB As typHLS) As compareResult
  If colorA.hue < colorB.hue Then
    compare = greater
  ElseIf colorA.hue > colorB.hue Then
    compare = less
  Else
    If colorA.saturation > colorB.saturation Then
      compare = greater
    ElseIf colorA.saturation < colorB.saturation Then
      compare = less
    Else
      If colorA.luminance < colorB.luminance Then
        compare = greater
      ElseIf colorA.luminance > colorB.luminance Then
        compare = less
      Else
        compare = equal
      End If
    End If
  End If
End Function
