Attribute VB_Name = "mCommon"
Option Explicit

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long

Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158

'/////////////////////////////////////////////////////////////////////////////////////////////////////
'## TextBox routines
'
Public Sub HiLite(txtBox As TextBox)
    With txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
'        .SetFocus
    End With
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////////
'## ComboBox routines
'
Public Function FindComboText(oCombo As ComboBox, ByVal Text As String) As Boolean

    Dim lResult As Long
    
    lResult = SendMessage(oCombo.hwnd, CB_FINDSTRING, 0&, ByVal (Text))
    If lResult Then
        oCombo.ListIndex = lResult
        FindComboText = True
    End If

End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////
'## Array routines
'
Public Function IsInArray(ByVal FindValue As Variant, _
                          ByVal arrSearch As Variant) As Boolean
    '@@ Original code by Brian Gillham
    On Error GoTo LocalError
    If Not IsArray(arrSearch) Then Exit Function
    If Not IsNumeric(FindValue) Then FindValue = UCase(FindValue)
    IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0
    Exit Function
LocalError:
    '## just in case...
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////
'## String routines
'
Public Function CompactSpaces(ByVal Text As String) As String
    '@@ Original code by Brian Gillham
    Dim sResult As String

    sResult = Trim$(Text)
    While InStr(sResult, String(2, " ")) > 0
        sResult = Replace(sResult, String(2, " "), " ")
    Wend
    CompactSpaces = sResult

End Function
