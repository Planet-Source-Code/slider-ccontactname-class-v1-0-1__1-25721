VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPersonName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================================
'
' Class Name:   cPersonName
' Author:       Slider
' Date:         01/08/2001
' Version:      01.00.00
' Description:  Handles the creation plus breaking down of a person's name
'               with auto-capitalise (option) and criteria-based validation.
' Edit History: 01.00.00 01/08/01 Initial Release
'
'===========================================================================

Option Explicit

Private msFullName    As String

Private msTitle       As String
Private msFirst       As String
Private msMiddle      As String
Private msLast        As String
Private msSuffix      As String

Private msTitles()    As String
Private msSuffixes()  As String

Private mbAutoCorrect As Boolean

Public Enum eComboType
    ectTitle = 0
    ectSuffix = 1
End Enum

Public Enum eValidateFields
    evfTitle = 1
    evfFirst = 2
    evfMiddle = 4
    evfLast = 8
    evfSuffix = 16
End Enum

'/////////////////////////////////////////////////////////////////////////////////////////////////
'
Public Property Get FullName() As String
    FullName = pBuildName
End Property

Public Property Let FullName(ByVal vNewValue As String)
    msFullName = CompactSpaces(vNewValue)
    If Len(msFullName) Then pExtractName
End Property

Public Property Get First() As String
    First = msFirst
End Property

Public Property Let First(ByVal vNewValue As String)
    msFirst = IIf(mbAutoCorrect, pProperNameCase(vNewValue), vNewValue)
End Property

Public Property Get Middle() As String
    Middle = msMiddle
End Property

Public Property Let Middle(ByVal vNewValue As String)
    msMiddle = IIf(mbAutoCorrect, pProperNameCase(vNewValue), vNewValue)
End Property

Public Property Get Last() As String
    Last = msLast
End Property

Public Property Let Last(ByVal vNewValue As String)
    msLast = IIf(mbAutoCorrect, pProperNameCase(vNewValue), vNewValue)
End Property

Public Property Get Title() As String
    Title = msTitle
End Property

Public Property Let Title(ByVal vNewValue As String)

    Dim lLoop As Long
    Dim bFnd  As Boolean

    For lLoop = 1 To UBound(msTitles)
        If UCase$(vNewValue) = UCase$(msTitles(lLoop)) Then
            If mbAutoCorrect Then
                msTitle = msTitles(lLoop)
            Else
                msTitle = vNewValue
            End If
            bFnd = True
            Exit For
        End If
    Next
    If Not bFnd Then
        msTitle = vNewValue
    End If

End Property

Public Property Get Suffix() As String
    Suffix = msSuffix
End Property

Public Property Let Suffix(ByVal vNewValue As String)

    Dim lLoop As Long
    Dim bFnd  As Boolean

    For lLoop = 1 To UBound(msSuffixes)
        If UCase$(vNewValue) = UCase$(msSuffixes(lLoop)) Then
            If mbAutoCorrect Then
                msSuffix = msSuffixes(lLoop)
            Else
                msSuffix = vNewValue
            End If
            bFnd = True
            Exit For
        End If
    Next
    If Not bFnd Then
        msSuffix = vNewValue
    End If

End Property

Public Property Get AutoCorrect() As Boolean
    AutoCorrect = mbAutoCorrect
End Property

Public Property Let AutoCorrect(ByVal State As Boolean)
    mbAutoCorrect = State
End Property

Public Sub FillComboBox(ByRef oCombo As VB.ComboBox, _
                        ByVal ComboType As eComboType)

    Dim lLoop As Long
    Dim lMax  As Long

    With oCombo
        .Clear
        Select Case ComboType
            Case ectTitle
                lMax = UBound(msTitles)
                For lLoop = 1 To lMax
                    .AddItem msTitles(lLoop)
                    .ItemData(.NewIndex) = lLoop
                Next

            Case ectSuffix: lMax = UBound(msSuffixes)
                lMax = UBound(msSuffixes)
                For lLoop = 1 To lMax
                    .AddItem msSuffixes(lLoop)
                    .ItemData(.NewIndex) = lLoop
                Next

        End Select
        If .ListCount Then
            .ListIndex = 0
        End If
    End With

End Sub

Public Function ValidateName(ByVal FullName As String, _
                             ByVal Fields As eValidateFields) As eValidateFields

    Dim stFullName As String
    Dim stTitle    As String
    Dim stFirst    As String
    Dim stMiddle   As String
    Dim stLast     As String
    Dim stSuffix   As String

    Dim eResult  As eValidateFields '## If valid then 'eResult' should equal 'Fields'.
                                    '   Unset result bits represent missing fields.

    If Len(FullName) Then

        '## Save fields
        stFullName = msFullName
        stTitle = msTitle
        stFirst = msFirst
        stMiddle = msMiddle
        stLast = msLast
        stSuffix = msSuffix

        msFullName = FullName
        pExtractName
        '## Calculate result code
        eResult = (Fields And evfTitle) * Abs(Len(msTitle) > 0) + _
                  (Fields And evfFirst) * Abs(Len(msFirst) > 0) + _
                  (Fields And evfMiddle) * Abs(Len(msMiddle) > 0) + _
                  (Fields And evfLast) * Abs(Len(msLast) > 0) + _
                  (Fields And evfSuffix) * Abs(Len(msSuffix) > 0)
        ValidateName = eResult

        '## Restore fields
        msFullName = stFullName
        msTitle = stTitle
        msFirst = stFirst
        msMiddle = stMiddle
        msLast = stLast
        msSuffix = stSuffix

    End If

End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////
'
Private Sub Class_Initialize()

    ReDim msTitles(1 To 6)
    msTitles(1) = "Dr."
    msTitles(2) = "Miss"
    msTitles(3) = "Mr."
    msTitles(4) = "Mrs."
    msTitles(5) = "Ms."
    msTitles(6) = "Prof."

    ReDim msSuffixes(1 To 5)
    msSuffixes(1) = "I"
    msSuffixes(2) = "II"
    msSuffixes(3) = "III"
    msSuffixes(4) = "Jr."
    msSuffixes(5) = "Sr."

End Sub

Private Sub Class_Terminate()
'
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////
'
Private Function pBuildName() As String

    Dim sFull As String

    sFull = Trim$(IIf(Len(msTitle), msTitle + " ", "") + _
                  IIf(Len(msFirst), msFirst + " ", "") + _
                  IIf(Len(msMiddle), msMiddle + " ", "") + _
                  IIf(Len(msLast), msLast + " ", "") + _
                  IIf(Len(msSuffix), msSuffix, ""))
    pBuildName = sFull

End Function

Private Sub pExtractName()

    Dim lCount    As Long
    Dim lMin      As Long
    Dim lMax      As Long
    Dim lLoop     As Long
    Dim sTokens() As String
    '
    '## Clear Fields
    '
    msFirst = ""
    msMiddle = ""
    msLast = ""
    msTitle = ""
    msSuffix = ""

    sTokens = Split(msFullName, " ")
    lCount = UBound(sTokens)
    lMax = lCount

    '## Check for Title
    For lLoop = 1 To UBound(msTitles)
        If UCase$(sTokens(0)) = UCase$(msTitles(lLoop)) Then
            If mbAutoCorrect Then
                msTitle = msTitles(lLoop)
            Else
                msTitle = sTokens(0)
            End If
            lMin = 1
            Exit For
        End If
    Next

    '## Check for suffix
    For lLoop = 1 To UBound(msSuffixes)
        If UCase$(sTokens(lCount)) = UCase$(msSuffixes(lLoop)) Then
            If mbAutoCorrect Then
                msSuffix = msSuffixes(lLoop)
            Else
                msSuffix = sTokens(lCount)
            End If
            lCount = lCount - 1
            lMax = lCount
            Exit For
        End If
    Next
    lCount = lCount - lMin

    '## Clean up rest of the name
    If mbAutoCorrect Then
        For lLoop = lMin To lMax
            sTokens(lLoop) = pProperNameCase(sTokens(lLoop))
        Next
    End If

    '## Break out remaining name tokens
    Select Case lCount
        Case 0                          '## Last name only
            msLast = sTokens(lMin)
        Case 1                          '## First & last names only
            msFirst = sTokens(lMin)
            msLast = sTokens(lMin + 1)
        Case Is > 1                     '## Handles multiple middle names
            msFirst = sTokens(lMin)
            msLast = sTokens(lMax)
            For lLoop = lMin + 1 To lMax - 1
                msMiddle = msMiddle + " " + sTokens(lLoop)
            Next
            msMiddle = Trim$(msMiddle)
    End Select

End Sub

'
'## First character capitalised, remaining lowercase except hypen'd words, 'Mc' & 'Mac'
'
Private Function pProperNameCase(ByVal Word As String) As String

    Dim sResult As String
    Dim lSep    As Long
    Dim lLen    As Long
    Dim lPos    As Long

    lLen = Len(Word)
    sResult = UCase$(Left$(Word, 1)) + LCase$(Mid$(Word, 2))
    lSep = InStr(Word, "-")
    If lSep And lSep < lLen Then Mid$(sResult, lSep + 1) = UCase$(Mid$(sResult, lSep + 1, 1))

    lPos = 1
    Do
        lSep = InStr(lPos, UCase$(Word), "'")
        If lSep Then
            lPos = lSep + 1
            If lSep + 1 < lLen Then
                Mid$(sResult, lSep, 1) = UCase$(Mid$(sResult, lSep, 1))
                Mid$(sResult, lSep + 1, 1) = UCase$(Mid$(sResult, lSep + 1, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    lPos = 1
    Do
        lSep = InStr(lPos, UCase$(Word), "MC")
        If lSep Then
            lPos = lSep + 1
            If lSep + 1 < lLen Then
                Mid$(sResult, lSep, 1) = UCase$(Mid$(sResult, lSep, 1))
                Mid$(sResult, lSep + 2, 1) = UCase$(Mid$(sResult, lSep + 2, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    lPos = 1
    Do
        lSep = InStr(lPos, UCase$(Word), "MAC")
        If lSep Then
            lPos = lSep + 1
            If lSep + 1 < lLen Then
                Mid$(sResult, lSep, 1) = UCase$(Mid$(sResult, lSep, 1))
                Mid$(sResult, lSep + 3, 1) = UCase$(Mid$(sResult, lSep + 3, 1))
            End If
        Else
            Exit Do
        End If
    Loop

    pProperNameCase = sResult

End Function
