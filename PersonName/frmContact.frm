VERSION 5.00
Begin VB.Form frmContact 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Test: cPersonName Class"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Close"
      Height          =   330
      Left            =   2835
      TabIndex        =   14
      Top             =   3990
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name &Details:"
      Height          =   2955
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   3900
      Begin VB.TextBox txtDialog 
         Height          =   285
         Index           =   1
         Left            =   1395
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   630
         Width           =   2325
      End
      Begin VB.TextBox txtDialog 
         Height          =   285
         Index           =   2
         Left            =   1395
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   945
         Width           =   2325
      End
      Begin VB.TextBox txtDialog 
         Height          =   285
         Index           =   3
         Left            =   1395
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1260
         Width           =   2325
      End
      Begin VB.ComboBox cboDialog 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   315
         Width           =   2325
      End
      Begin VB.ComboBox cboDialog 
         Height          =   315
         Index           =   1
         Left            =   1395
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1575
         Width           =   2325
      End
      Begin VB.CheckBox chkDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Auto Capitalise:"
         Height          =   330
         Left            =   105
         TabIndex        =   13
         Top             =   1890
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Title: "
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "First &Name: "
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Middle Name: "
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   7
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Last Name: "
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   1260
         Width           =   1170
      End
      Begin VB.Label lblDialog 
         Alignment       =   1  'Right Justify
         Caption         =   "&Suffix: "
         Height          =   225
         Index           =   5
         Left            =   135
         TabIndex        =   11
         Top             =   1575
         Width           =   1170
      End
   End
   Begin VB.TextBox txtDialog 
      Height          =   285
      Index           =   0
      Left            =   1365
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   210
      Width           =   2640
   End
   Begin VB.Label lblDialog 
      Alignment       =   1  'Right Justify
      Caption         =   "&FullName: "
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1170
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    frmContact
' Author:       Slider
' Date:         01/08/2001
' Version:      01.00.00
' Description:  Test form for cPersonName Class.
' Edit History: 01.00.00 01/08/01 Initial Release
'               01.00.01 04/08/01 Fixed Issue with cboDialog_Validate &
'                                 Auto-complete does not set ListIndex.
' Notes:        This test app was designed to exploit all of the features
'               available in the cPersonName Class such as:-
'                   * Convert Fullname to 5 individual fields
'                   * Convert 5 individual fields to Fullname
'                   * Automatically capitalise fields/fullname (option)
'                   * Validate the fullname based on set criteria
'               The test apps also illustrates (for beginners) how to:-
'                   * Auto-complete a ComboBox
'                   * Quick ComboBox search using API
'                   * Simple field hilighting methods for TextBox and
'                     ComboBox
'                   * Avoid complex If/Then structures using bitwise
'                     operation and the IIF function
'                   * Avoid infinite event loops (Stack overflow errors)
'                   * Encapsulating data and associated functions into a
'                     reusable code class
'
'===========================================================================

Option Explicit

Private Enum eTextBox
    etbFullName = 0
    etbFirst = 1
    etbMiddle = 2
    etbLast = 3
End Enum

Private Enum eComboBox
    ecbTitle = 0
    ecbSuffix = 1
End Enum

Private mcPersonName   As cPersonName

'Private mbCboLoading       As Boolean
Private mbCboExist(0 To 1) As Boolean
Private mbBackspaced       As Boolean
Private mbIsDirty          As Boolean
Private mbIsBusy           As Boolean

'/////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cboDialog_Change(Index As Integer)

    If mbIsBusy Then Exit Sub
'    If mbCboLoading Then Exit Sub

    mbIsDirty = True
    '## Set Search Toolbat Button state
    If Len(cboDialog(Index).Text) > 0 Then
        mbCboExist(Index) = True
    Else
        mbCboExist(Index) = False
    End If

    '## Auto-complete combobox
    '## If firing in response to a backspace or delete, don't run the auto-complete
    '   complete code. (Otherwise you wouldn't be able to back up.)
    If mbBackspaced = True Or cboDialog(Index).Text = "" Then
        mbBackspaced = False
        Exit Sub
    End If

    Dim lLoop As Long
    Dim nSel  As Long

    '## Run through the available items and grab the first matching one.
    For lLoop = 0 To cboDialog(Index).ListCount - 1
        If InStr(1, cboDialog(Index).List(lLoop), cboDialog(Index).Text, vbTextCompare) = 1 Then
            '## Save the SelStart property.
            nSel = cboDialog(Index).SelStart
            cboDialog(Index).Text = cboDialog(Index).List(lLoop)
            '## Set the selection in the combo.
            cboDialog(Index).SelStart = nSel
            cboDialog(Index).SelLength = Len(cboDialog(Index).Text) - nSel
            Exit For
        End If
    Next

End Sub

Private Sub cboDialog_Click(Index As Integer)
    If mbIsBusy Then Exit Sub
'    If mbCboLoading Then Exit Sub
    mbCboExist(Index) = True
    mbIsDirty = True            '## A change was made...
End Sub

Private Sub cboDialog_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    '## Auto-complete combobox
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If cboDialog(Index).Text <> "" Then
            '## Let the Change event know that it shouldn't respond to this change.
            mbBackspaced = True
        End If
    End If

End Sub

Private Sub cboDialog_KeyPress(Index As Integer, KeyAscii As Integer)
    Debug.Print KeyAscii
    If KeyAscii = 13 Then
        If mbCboExist(Index) Then
            '## Special code event...
        End If
    End If
End Sub

Private Sub cboDialog_Validate(Index As Integer, Cancel As Boolean)

    '## We're leaving this field...
    If mbIsDirty Then           '## Anything changed?
        mbIsBusy = True
        With cboDialog(Index)
            Select Case Index
                Case ecbTitle:  mcPersonName.Title = .Text
                Case ecbSuffix: mcPersonName.Suffix = .Text
            End Select
            FindComboText cboDialog(Index), .Text
            txtDialog(etbFullName).Text = mcPersonName.FullName
        End With
        mbIsBusy = False
    End If
    mbIsDirty = False

End Sub

Private Sub chkDialog_Click()
    mcPersonName.AutoCorrect = CBool(chkDialog.Value)
End Sub

Private Sub cmdDialog_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set mcPersonName = New cPersonName
    mbIsBusy = True             '## disable the need to respond to specific VB events
    With mcPersonName
        .AutoCorrect = CBool(chkDialog.Value)
        .Title = "Mr."
        .First = "Ted"
        .Middle = "De"
        .Last = "Bair"
        .Suffix = "Sr."
        .FillComboBox cboDialog(ecbTitle), ectTitle
        .FillComboBox cboDialog(ecbSuffix), ectSuffix
        '
        '## Fill GUI Fields with data
        '
        mbCboExist(ecbTitle) = FindComboText(cboDialog(ecbTitle), .Title)
        mbCboExist(ecbSuffix) = FindComboText(cboDialog(ecbSuffix), .Suffix)
        txtDialog(etbFullName).Text = .FullName
        txtDialog(etbFirst).Text = .First
        txtDialog(etbMiddle).Text = .Middle
        txtDialog(etbLast).Text = .Last
    End With
    mbIsBusy = False            '## re-enable the app's VB event handling

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcPersonName = Nothing
End Sub

Private Sub txtDialog_Change(Index As Integer)
'    Debug.Print "txtDialog_Change"
    If mbIsBusy Then Exit Sub
    mbIsDirty = True            '## A change was made...
End Sub

Private Sub txtDialog_GotFocus(Index As Integer)
    HiLite txtDialog(Index)
End Sub

Private Sub txtDialog_Validate(Index As Integer, Cancel As Boolean)

'    Debug.Print "txtDialog_Validate"
    '## We're leaving this field...
    If mbIsDirty Then           '## Anything changed?
        mbIsBusy = True         '## disable the need to respond to specific VB events
        With mcPersonName
            Select Case Index
                Case etbFullName
                    Dim eTest   As eValidateFields
                    Dim eResult As eValidateFields

                    .FullName = txtDialog(etbFullName).Text
                    If CBool(chkDialog.Value) Then
                        txtDialog(etbFullName).Text = .FullName
                    End If
                    txtDialog(etbFirst).Text = .First   '## returns extracted fields
                    txtDialog(etbMiddle).Text = .Middle '
                    txtDialog(etbLast).Text = .Last     '
                    mbCboExist(ecbTitle) = FindComboText(cboDialog(ecbTitle), .Title)
                    mbCboExist(ecbSuffix) = FindComboText(cboDialog(ecbSuffix), .Suffix)
                    '
                    '## Test if specific data was entered...
                    '
                    eTest = evfFirst + evfMiddle + evfLast
                    eResult = .ValidateName(txtDialog(etbFullName).Text, eTest)
                    If eResult <> eTest Then
                        MsgBox "Incomplete name. The following fields were missing:" + vbCrLf + vbCrLf + _
                                   IIf((eResult And evfFirst), "", vbTab + "First/Christian name" + vbCrLf) + _
                                   IIf((eResult And evfMiddle), "", vbTab + "Middle name(s)" + vbCrLf) + _
                                   IIf((eResult And evfLast), "", vbTab + "Last/Given name" + vbCrLf), _
                                vbInformation + vbOKOnly, _
                                "WARNING!"
                    End If

                Case etbFirst
                    .First = txtDialog(etbFirst).Text
                    txtDialog(etbFirst).Text = .First       '## reformats keyed field
                    txtDialog(etbFullName).Text = .FullName '## returns formatted name

                Case etbMiddle
                    .Middle = txtDialog(etbMiddle).Text
                    txtDialog(etbMiddle).Text = .Middle     '## reformats keyed field
                    txtDialog(etbFullName).Text = .FullName '## returns formatted name

                Case etbLast
                    .Last = txtDialog(etbLast).Text
                    txtDialog(etbLast).Text = .Last         '## reformats keyed field
                    txtDialog(etbFullName).Text = .FullName '## returns formatted name

            End Select
        End With
        mbIsBusy = False    '## re-enable the app's VB event handling
    End If
    mbIsDirty = False       '## Changes applied, therefore reset dirty flag

End Sub
