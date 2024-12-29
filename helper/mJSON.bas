Attribute VB_Name = "mJSON"
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Description: copied from https://github.com/VBA-tools/VBA-Web/JSON converter  |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/
Dim oLogger    As iLogger
Public options As jsonOptions

Private Type jsonOptions
  ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
  ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
  ' See: http://support.microsoft.com/kb/269370
  '
  ' By default, mJSON will use String for numbers longer than 15 characters that contain only digits
  ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
  useDoubleForLargeNumbers As Boolean

  ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
  allowUnquotedKeys As Boolean

  ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in generate
  escapeSolidus As Boolean
End Type

Private Property Get logger() As iLogger
  If oLogger Is Nothing Then Set oLogger = config.logger
  Set logger = oLogger
End Property

' Convert object (Dictionary/Collection/Array) to JSON
' @method generate
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
Public Function generate(ByVal jsonValue As Variant, _
                         Optional ByVal whiteSpace As Variant, _
                         Optional ByVal json_currentIndentation As Long = 0) As String
  Dim json_buffer           As String
  Dim json_bufferPosition   As Long
  Dim json_bufferLength     As Long
  Dim json_index            As Long
  Dim json_lBound           As Long
  Dim json_uBound           As Long
  Dim json_isFirstItem      As Boolean
  Dim json_index2D          As Long
  Dim json_lBound2D         As Long
  Dim json_uBound2D         As Long
  Dim json_isFirstItem2D    As Boolean
  Dim json_key              As Variant
  Dim json_value            As Variant
  Dim json_dateStr          As String
  Dim json_converted        As String
  Dim json_skipItem         As Boolean
  Dim json_prettyPrint      As Boolean
  Dim json_indentation      As String
  Dim json_innerIndentation As String

  json_lBound = -1
  json_uBound = -1
  json_isFirstItem = True
  json_lBound2D = -1
  json_uBound2D = -1
  json_isFirstItem2D = True
  json_prettyPrint = Not IsMissing(whiteSpace)

  Select Case VBA.VarType(jsonValue)
  Case VBA.vbNull
    generate = "null"
  Case VBA.vbDate
    ' Date
    json_dateStr = mIso8601.generate(VBA.CDate(jsonValue))

    generate = """" & json_dateStr & """"
  Case VBA.vbString
    ' String (or large number encoded as string)
    If Not options.useDoubleForLargeNumbers And stringIsLargeNumber(jsonValue) Then
      generate = jsonValue
    Else
      generate = """" & encode(jsonValue) & """"
    End If
  Case VBA.vbBoolean
    If jsonValue Then
      generate = "true"
    Else
      generate = "false"
    End If
  Case VBA.vbArray To VBA.vbArray + VBA.vbByte
    If json_prettyPrint Then
      If VBA.VarType(whiteSpace) = VBA.vbString Then
        json_indentation = VBA.String$(json_currentIndentation + 1, whiteSpace)
        json_innerIndentation = VBA.String$(json_currentIndentation + 2, whiteSpace)
      Else
        json_indentation = VBA.Space$((json_currentIndentation + 1) * whiteSpace)
        json_innerIndentation = VBA.Space$((json_currentIndentation + 2) * whiteSpace)
      End If
    End If

    ' Array
    bufferAppend json_buffer, "[", json_bufferPosition, json_bufferLength

    On Error Resume Next

    json_lBound = LBound(jsonValue, 1)
    json_uBound = UBound(jsonValue, 1)
    json_lBound2D = LBound(jsonValue, 2)
    json_uBound2D = UBound(jsonValue, 2)

    If json_lBound >= 0 And json_uBound >= 0 Then
      For json_index = json_lBound To json_uBound
        If json_isFirstItem Then
          json_isFirstItem = False
        Else
          ' Append comma to previous line
          bufferAppend json_buffer, ",", json_bufferPosition, json_bufferLength
        End If

        If json_lBound2D >= 0 And json_uBound2D >= 0 Then
          ' 2D Array
          If json_prettyPrint Then
            bufferAppend json_buffer, vbNewLine, json_bufferPosition, json_bufferLength
          End If
          bufferAppend json_buffer, json_indentation & "[", json_bufferPosition, json_bufferLength

          For json_index2D = json_lBound2D To json_uBound2D
            If json_isFirstItem2D Then
              json_isFirstItem2D = False
            Else
              bufferAppend json_buffer, ",", json_bufferPosition, json_bufferLength
            End If

            json_converted = generate(jsonValue(json_index, json_index2D), whiteSpace, json_currentIndentation + 2)

            ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
            If json_converted = "" Then
              ' (nest to only check if converted = "")
              If isUndefined(jsonValue(json_index, json_index2D)) Then
                  json_converted = "null"
              End If
            End If

            If json_prettyPrint Then
              json_converted = vbNewLine & json_innerIndentation & json_converted
            End If

            bufferAppend json_buffer, json_converted, json_bufferPosition, json_bufferLength
          Next json_index2D
          If json_prettyPrint Then
            bufferAppend json_buffer, vbNewLine, json_bufferPosition, json_bufferLength
          End If

          bufferAppend json_buffer, json_indentation & "]", json_bufferPosition, json_bufferLength
          json_isFirstItem2D = True
        Else
          ' 1D Array
          json_converted = generate(jsonValue(json_index), whiteSpace, json_currentIndentation + 1)

          ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
          If json_converted = "" Then
            ' (nest to only check if converted = "")
            If isUndefined(jsonValue(json_index)) Then
              json_converted = "null"
            End If
          End If

          If json_prettyPrint Then
            json_converted = vbNewLine & json_indentation & json_converted
          End If

          bufferAppend json_buffer, json_converted, json_bufferPosition, json_bufferLength
        End If
      Next json_index
    End If

    On Error GoTo 0

    If json_prettyPrint Then
      bufferAppend json_buffer, vbNewLine, json_bufferPosition, json_bufferLength

      If VBA.VarType(whiteSpace) = VBA.vbString Then
        json_indentation = VBA.String$(json_currentIndentation, whiteSpace)
      Else
        json_indentation = VBA.Space$(json_currentIndentation * whiteSpace)
      End If
    End If

    bufferAppend json_buffer, json_indentation & "]", json_bufferPosition, json_bufferLength

    generate = bufferToString(json_buffer, json_bufferPosition)

  ' Dictionary or Collection
  Case VBA.vbObject
    If json_prettyPrint Then
      If VBA.VarType(whiteSpace) = VBA.vbString Then
        json_indentation = VBA.String$(json_currentIndentation + 1, whiteSpace)
      Else
        json_indentation = VBA.Space$((json_currentIndentation + 1) * whiteSpace)
      End If
    End If

    ' Dictionary
    If VBA.TypeName(jsonValue) = "Dictionary" Then
      bufferAppend json_buffer, "{", json_bufferPosition, json_bufferLength
      For Each json_key In jsonValue.Keys
        ' For Objects, undefined (Empty/Nothing) is not added to object
        json_converted = generate(jsonValue(json_key), whiteSpace, json_currentIndentation + 1)
        If json_converted = "" Then
          json_skipItem = isUndefined(jsonValue(json_key))
        Else
          json_skipItem = False
        End If

        If Not json_skipItem Then
          If json_isFirstItem Then
            json_isFirstItem = False
          Else
            bufferAppend json_buffer, ",", json_bufferPosition, json_bufferLength
          End If

          If json_prettyPrint Then
            json_converted = vbNewLine & json_indentation & """" & json_key & """: " & json_converted
          Else
            json_converted = """" & json_key & """:" & json_converted
          End If

          bufferAppend json_buffer, json_converted, json_bufferPosition, json_bufferLength
        End If
      Next json_key

      If json_prettyPrint Then
        bufferAppend json_buffer, vbNewLine, json_bufferPosition, json_bufferLength

        If VBA.VarType(whiteSpace) = VBA.vbString Then
          json_indentation = VBA.String$(json_currentIndentation, whiteSpace)
        Else
          json_indentation = VBA.Space$(json_currentIndentation * whiteSpace)
        End If
      End If

      bufferAppend json_buffer, json_indentation & "}", json_bufferPosition, json_bufferLength

    ' Collection
    ElseIf VBA.TypeName(jsonValue) = "Collection" Then
      bufferAppend json_buffer, "[", json_bufferPosition, json_bufferLength
      For Each json_value In jsonValue
        If json_isFirstItem Then
          json_isFirstItem = False
        Else
          bufferAppend json_buffer, ",", json_bufferPosition, json_bufferLength
        End If

        json_converted = generate(json_value, whiteSpace, json_currentIndentation + 1)

        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
        If json_converted = "" Then
          ' (nest to only check if converted = "")
          If isUndefined(json_value) Then
            json_converted = "null"
          End If
        End If

        If json_prettyPrint Then
          json_converted = vbNewLine & json_indentation & json_converted
        End If

        bufferAppend json_buffer, json_converted, json_bufferPosition, json_bufferLength
      Next json_value

      If json_prettyPrint Then
        bufferAppend json_buffer, vbNewLine, json_bufferPosition, json_bufferLength

        If VBA.VarType(whiteSpace) = VBA.vbString Then
          json_indentation = VBA.String$(json_currentIndentation, whiteSpace)
        Else
          json_indentation = VBA.Space$(json_currentIndentation * whiteSpace)
        End If
      End If

      bufferAppend json_buffer, json_indentation & "]", json_bufferPosition, json_bufferLength
    End If

    generate = bufferToString(json_buffer, json_bufferPosition)
  Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
    ' Number (use decimals for numbers)
    generate = VBA.Replace(jsonValue, ",", ".")
  Case Else
    ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
    ' Use VBA's built-in to-string
    On Error Resume Next
    generate = jsonValue
    On Error GoTo 0
  End Select
End Function

Private Function parseObject(json_string As String, ByRef json_index As Long) As Object
  Dim json_key As String
  Dim json_nextChar As String

  Set parseObject = CreateObject("Scripting.Dictionary")
  SkipSpaces json_string, json_index
  If VBA.Mid$(json_string, json_index, 1) <> "{" Then
    Err.Raise -10001, "mJSON.parseObject", parseErrorMessage(json_string, json_index, "Expecting '{'")
  Else
    json_index = json_index + 1

    Do
      SkipSpaces json_string, json_index
      If VBA.Mid$(json_string, json_index, 1) = "}" Then
        json_index = json_index + 1
        Exit Function
      ElseIf VBA.Mid$(json_string, json_index, 1) = "," Then
        json_index = json_index + 1
        SkipSpaces json_string, json_index
      End If

      json_key = parseKey(json_string, json_index)
      json_nextChar = peek(json_string, json_index)
      If json_nextChar = "[" Or json_nextChar = "{" Then
        Set parseObject.Item(json_key) = parseValue(json_string, json_index)
      Else
        parseObject.Item(json_key) = parseValue(json_string, json_index)
      End If
    Loop
  End If
End Function

Private Function parseArray(json_string As String, ByRef json_index As Long) As Collection
  Set parseArray = New Collection

  Call SkipSpaces(json_string, json_index)
  If VBA.Mid$(json_string, json_index, 1) <> "[" Then
    Err.Raise -10001, "mJSON.parseArray", parseErrorMessage(json_string, json_index, "Expecting '['")
  Else
    json_index = json_index + 1

    Do
      Call SkipSpaces(json_string, json_index)
      If VBA.Mid$(json_string, json_index, 1) = "]" Then
        json_index = json_index + 1
        Exit Function
      ElseIf VBA.Mid$(json_string, json_index, 1) = "," Then
        json_index = json_index + 1
        SkipSpaces json_string, json_index
      End If

      parseArray.Add parseValue(json_string, json_index)
    Loop
  End If
End Function

Private Function parseValue(json_string As String, ByRef json_index As Long) As Variant
  Call SkipSpaces(json_string, json_index)
  Select Case VBA.Mid$(json_string, json_index, 1)
  Case "{"
    Set parseValue = parseObject(json_string, json_index)
  Case "["
    Set parseValue = parseArray(json_string, json_index)
  Case """", "'"
    parseValue = parseString(json_string, json_index)
  Case Else
    If VBA.Mid$(json_string, json_index, 4) = "true" Then
      parseValue = True
      json_index = json_index + 4
    ElseIf VBA.Mid$(json_string, json_index, 5) = "false" Then
      parseValue = False
      json_index = json_index + 5
    ElseIf VBA.Mid$(json_string, json_index, 4) = "null" Then
      parseValue = Null
      json_index = json_index + 4
    ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_string, json_index, 1)) Then
      parseValue = parseNumber(json_string, json_index)
    Else
      Err.Raise -10001, "mJSON.parseValue", parseErrorMessage(json_string, json_index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
    End If
  End Select
End Function

Private Function parseString(json_string As String, ByRef json_index As Long) As String
  Dim json_quote          As String
  Dim json_char           As String
  Dim json_code           As String
  Dim json_buffer         As String
  Dim json_bufferPosition As Long
  Dim json_bufferLength   As Long

  Call SkipSpaces(json_string, json_index)

  ' Store opening quote to look for matching closing quote
  json_quote = VBA.Mid$(json_string, json_index, 1)
  json_index = json_index + 1

  Do While json_index > 0 And json_index <= Len(json_string)
    json_char = VBA.Mid$(json_string, json_index, 1)

    Select Case json_char
    Case "\"
      ' Escaped string, \\, or \/
      json_index = json_index + 1
      json_char = VBA.Mid$(json_string, json_index, 1)

      Select Case json_char
      Case """", "\", "/", "'"
        bufferAppend json_buffer, json_char, json_bufferPosition, json_bufferLength
        json_index = json_index + 1
      Case "b"
        bufferAppend json_buffer, vbBack, json_bufferPosition, json_bufferLength
        json_index = json_index + 1
      Case "f"
        bufferAppend json_buffer, vbFormFeed, json_bufferPosition, json_bufferLength
        json_index = json_index + 1
      Case "n"
        bufferAppend json_buffer, vbCrLf, json_bufferPosition, json_bufferLength
        json_index = json_index + 1
      Case "r"
        bufferAppend json_buffer, vbCr, json_bufferPosition, json_bufferLength
        json_index = json_index + 1
      Case "t"
        bufferAppend json_buffer, vbTab, json_bufferPosition, json_bufferLength
        json_index = json_index + 1
      Case "u"
        ' Unicode character escape (e.g. \u00a9 = Copyright)
        json_index = json_index + 1
        json_code = VBA.Mid$(json_string, json_index, 4)
        bufferAppend json_buffer, VBA.ChrW(VBA.Val("&h" + json_code)), json_bufferPosition, json_bufferLength
        json_index = json_index + 4
      End Select
    Case json_quote
      parseString = bufferToString(json_buffer, json_bufferPosition)
      json_index = json_index + 1
      Exit Function
    Case Else
      bufferAppend json_buffer, json_char, json_bufferPosition, json_bufferLength
      json_index = json_index + 1
    End Select
  Loop
End Function

Private Function parseNumber(json_string As String, ByRef json_index As Long) As Variant
  Dim json_char          As String
  Dim json_value         As String
  Dim json_isLargeNumber As Boolean

  Call SkipSpaces(json_string, json_index)

  Do While json_index > 0 And json_index <= Len(json_string)
    json_char = VBA.Mid$(json_string, json_index, 1)

    If VBA.InStr("+-0123456789.eE", json_char) Then
      ' Unlikely to have massive number, so use simple append rather than buffer here
      json_value = json_value & json_char
      json_index = json_index + 1
    Else
      ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
      ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
      ' See: http://support.microsoft.com/kb/269370
      '
      ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
      ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
      json_isLargeNumber = IIf(InStr(json_value, "."), Len(json_value) >= 17, Len(json_value) >= 16)
      If Not options.useDoubleForLargeNumbers And json_isLargeNumber Then
        parseNumber = json_value
      Else
        ' VBA.Val does not use regional settings, so guard for comma is not needed
        parseNumber = VBA.Val(json_value)
      End If
      Exit Function
    End If
  Loop
End Function

Private Function parseKey(json_string As String, ByRef json_index As Long) As String
  ' Parse key with single or double quotes
  If VBA.Mid$(json_string, json_index, 1) = """" Or VBA.Mid$(json_string, json_index, 1) = "'" Then
    parseKey = parseString(json_string, json_index)
  ElseIf options.allowUnquotedKeys Then
    Dim json_char As String
    Do While json_index > 0 And json_index <= Len(json_string)
      json_char = VBA.Mid$(json_string, json_index, 1)
      If (json_char <> " ") And (json_char <> ":") Then
        parseKey = parseKey & json_char
        json_index = json_index + 1
      Else
        Exit Do
      End If
    Loop
  Else
    Err.Raise -10001, "mJSON.parseKey", parseErrorMessage(json_string, json_index, "Expecting '""' or '''")
  End If

  ' Check for colon and skip if present or throw if not present
  SkipSpaces json_string, json_index
  If VBA.Mid$(json_string, json_index, 1) <> ":" Then
    Err.Raise -10001, "mJSON.parseKey", parseErrorMessage(json_string, json_index, "Expecting ':'")
  Else
    json_index = json_index + 1
  End If
End Function

Private Function isUndefined(ByVal json_value As Variant) As Boolean
  ' Empty / Nothing -> undefined
  Select Case VBA.VarType(json_value)
  Case VBA.vbEmpty
    isUndefined = True
  Case VBA.vbObject
    Select Case VBA.TypeName(json_value)
    Case "Empty", "Nothing"
      isUndefined = True
    End Select
  End Select
End Function

Private Function encode(ByVal json_Text As Variant) As String
  ' Reference: http://www.ietf.org/rfc/rfc4627.txt
  ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
  Dim json_index          As Long
  Dim json_char           As String
  Dim json_ascCode        As Long
  Dim json_buffer         As String
  Dim json_bufferPosition As Long
  Dim json_bufferLength   As Long

  For json_index = 1 To VBA.Len(json_Text)
    json_char = VBA.Mid$(json_Text, json_index, 1)
    json_ascCode = VBA.AscW(json_char)

    ' When AscW returns a negative number, it returns the twos complement form of that number.
    ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
    ' https://support.microsoft.com/en-us/kb/272138
    If json_ascCode < 0 Then
      json_ascCode = json_ascCode + 65536
    End If

    ' From spec, ", \, and control characters must be escaped (solidus is optional)

    Select Case json_ascCode
    Case 34
      ' " -> 34 -> \"
      json_char = "\"""
    Case 92
      ' \ -> 92 -> \\
      json_char = "\\"
    Case 47
      ' / -> 47 -> \/ (optional)
      If options.escapeSolidus Then
        json_char = "\/"
      End If
    Case 8
      ' backspace -> 8 -> \b
      json_char = "\b"
    Case 12
      ' form feed -> 12 -> \f
      json_char = "\f"
    Case 10
      ' line feed -> 10 -> \n
      json_char = "\n"
    Case 13
      ' carriage return -> 13 -> \r
      json_char = "\r"
    Case 9
      ' tab -> 9 -> \t
      json_char = "\t"
    Case 0 To 31, 127 To 65535
      ' Non-ascii characters -> convert to 4-digit hex
      json_char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_ascCode), 4)
    End Select

    bufferAppend json_buffer, json_char, json_bufferPosition, json_bufferLength
  Next json_index
  encode = bufferToString(json_buffer, json_bufferPosition)
End Function

Private Function peek(json_string As String, ByVal json_index As Long, Optional json_numberOfCharacters As Long = 1) As String
  ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
  SkipSpaces json_string, json_index
  peek = VBA.Mid$(json_string, json_index, json_numberOfCharacters)
End Function

Private Sub SkipSpaces(json_string As String, ByRef json_index As Long)
  ' Increment index to skip over spaces
  Do While json_index > 0 And json_index <= VBA.Len(json_string) And VBA.Mid$(json_string, json_index, 1) = " "
    json_index = json_index + 1
  Loop
End Sub

Private Function stringIsLargeNumber(json_string As Variant) As Boolean
  ' Check if the given string is considered a "large number"
  ' (See parseNumber)

  Dim json_length As Long
  Dim json_charIndex As Long
  json_length = VBA.Len(json_string)

  ' Length with be at least 16 characters and assume will be less than 100 characters
  If json_length >= 16 And json_length <= 100 Then
    Dim json_CharCode As String

    stringIsLargeNumber = True

    For json_charIndex = 1 To json_length
      json_CharCode = VBA.Asc(VBA.Mid$(json_string, json_charIndex, 1))
      Select Case json_CharCode
      ' Look for .|0-9|E|e
      Case 46, 48 To 57, 69, 101
        ' Continue through characters
      Case Else
        stringIsLargeNumber = False
        Exit Function
      End Select
    Next json_charIndex
  End If
End Function

Private Function parseErrorMessage(json_string As String, ByRef json_index As Long, errorMessage As String) As String
  ' Provide detailed parse error message, including details of where and what occurred
  '
  ' Example:
  ' Error parsing JSON:
  ' {"abcde":True}
  '          ^
  ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

  Dim json_startIndex As Long
  Dim json_stopIndex As Long

  ' Include 10 characters before and after error (if possible)
  json_startIndex = json_index - 10
  json_stopIndex = json_index + 10
  If json_startIndex <= 0 Then json_startIndex = 1
  If json_stopIndex > VBA.Len(json_string) Then json_stopIndex = VBA.Len(json_string)

  logger.log FAILURE, "Error parsing JSON:" & VBA.vbNewLine & _
                      VBA.Mid$(json_string, json_startIndex, json_stopIndex - json_startIndex + 1) & VBA.vbNewLine & _
                      VBA.Space$(json_index - json_startIndex) & "^" & VBA.vbNewLine & _
                      errorMessage
  parseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                      VBA.Mid$(json_string, json_startIndex, json_stopIndex - json_startIndex + 1) & VBA.vbNewLine & _
                      errorMessage
  
End Function

Private Sub bufferAppend(ByRef json_buffer As String, _
                         ByRef json_append As Variant, _
                         ByRef json_bufferPosition As Long, _
                         ByRef json_bufferLength As Long)
  ' VBA can be slow to append strings due to allocating a new string for each append
  ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
  '
  ' Example:
  ' Buffer: "abc  "
  ' Append: "def"
  ' Buffer Position: 3
  ' Buffer Length: 5
  '
  ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
  ' Buffer: "abc       "
  ' Buffer Length: 10
  '
  ' Put "def" into buffer at position 3 (0-based)
  ' Buffer: "abcdef    "
  '
  ' Approach based on cStringBuilder from vbAccelerator
  ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
  '
  ' and clsStringAppend from Philip Swannell
  ' https://github.com/VBA-tools/VBA-JSON/pull/82

  Dim json_appendLength As Long
  Dim json_lengthPlusPosition As Long

  json_appendLength = VBA.Len(json_append)
  json_lengthPlusPosition = json_appendLength + json_bufferPosition

  If json_lengthPlusPosition > json_bufferLength Then
    ' Appending would overflow buffer, add chunk
    ' (double buffer length or append length, whichever is bigger)
    Dim json_AddedLength As Long
    json_AddedLength = IIf(json_appendLength > json_bufferLength, json_appendLength, json_bufferLength)

    json_buffer = json_buffer & VBA.Space$(json_AddedLength)
    json_bufferLength = json_bufferLength + json_AddedLength
  End If

  ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
  ' Function call on left-hand side of assignment must return Variant or Object
  Mid$(json_buffer, json_bufferPosition + 1, json_appendLength) = CStr(json_append)
  json_bufferPosition = json_bufferPosition + json_appendLength
End Sub

Private Function bufferToString(ByRef json_buffer As String, ByVal json_bufferPosition As Long) As String
  If json_bufferPosition > 0 Then bufferToString = VBA.Left$(json_buffer, json_bufferPosition)
End Function

' Convert JSON string to object (Dictionary/Collection)
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
Public Function parse(ByVal jsonString As String) As Object
  Dim json_index As Long
  json_index = 1
  logger.log DEBUGGER, "parse input json string"
  ' Remove vbCr, vbLf, and vbTab from json_String
  jsonString = VBA.Replace(VBA.Replace(VBA.Replace(jsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

  SkipSpaces jsonString, json_index
  Select Case VBA.Mid$(jsonString, json_index, 1)
  Case "{"
    Set parse = parseObject(jsonString, json_index)
  Case "["
    Set parse = parseArray(jsonString, json_index)
  Case Else
    Err.Raise -10001, "mJSON.parse", parseErrorMessage(jsonString, json_index, "Expecting '{' or '['")
  End Select
End Function
