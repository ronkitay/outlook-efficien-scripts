' ------------------------------------------------------------
' Name: StringUtils.vbs
' Description: Contains all functions related to String manipulation
' ------------------------------------------------------------
Attribute VB_Name = "StringUtils"

' ------------------------------------------------------------------------------------------------------------------------

' Name: StringStartsWith
' Description: Indicates if one string starts with another string
' Arguments:
'			StringToLookIn - The string to look in. E.g. "Hello World"
'			StringToLookFor - The string to look for. E.g. "Hell"
'			CompareType - (optional) by default 'vbBinaryCompare'. Use 'vbTextCompare' for non-case sensitive comparison.
' Returns: 
'		True - if StringToLookIn starts with StringToLookFor, false otherwise.
'		Case sensitive by default.  If you want non-case sensitive, set last parameter to vbTextCompare
' Examples:
'		MsgBox StringStartsWith("Test", "TE") = false
'		MsgBox StringStartsWith("Test", "TE", vbTextCompare) = true
Public Function StringStartsWith(ByVal StringToLookIn As String, StringToLookFor As String, Optional CompareType As VbCompareMethod = vbBinaryCompare) As Boolean
    
  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(StringToLookFor)
  If lLen > Len(StringToLookIn) Then Exit Function
  sCompare = Left(StringToLookIn, lLen)
  StringStartsWith = StrComp(sCompare, StringToLookFor, CompareType) = 0

End Function

' Name: StringEndsWith
' Description: Indicates if one string ends with another string
' Arguments:
'			StringToLookIn - The string to look in. E.g. "Hello World"
'			StringToLookFor - The string to look for. E.g. "world"
'			CompareType - (optional) by default 'vbBinaryCompare'. Use 'vbTextCompare' for non-case sensitive comparison.
' Returns: 
'		True - if StringToLookIn ends with StringToLookFor, false otherwise.
'		Case sensitive by default.  If you want non-case sensitive, set last parameter to vbTextCompare
' Examples:
'		MsgBox StringEndsWith("Test", "ST") = False
'		MsgBox StringEndsWith("Test", "ST", vbTextCompare) = True
Public Function StringEndsWith(ByVal StringToLookIn As String, StringToLookFor As String, Optional CompareType As VbCompareMethod = vbBinaryCompare) As Boolean

  Dim sCompare As String
  Dim lLen As Long
   
  lLen = Len(StringToLookFor)
  If lLen > Len(StringToLookIn) Then Exit Function
  sCompare = Right(StringToLookIn, lLen)
  StringEndsWith = StrComp(sCompare, StringToLookFor, CompareType) = 0

End Function

' ------------------------------------------------------------------------------------------------------------------------