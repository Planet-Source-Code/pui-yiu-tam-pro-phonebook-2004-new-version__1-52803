VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter"
   ClientHeight    =   2715
   ClientLeft      =   3300
   ClientTop       =   6255
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   6015
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter"
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cboField 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkMatch 
         Caption         =   "&Match whole word only"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter What:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter in Field:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   8880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmFilter.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Data - PhoneBook 2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image imgLogo 
      Height          =   1335
      Left            =   3720
      Picture         =   "frmFilter.frx":0442
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   3330
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit        'All variables that we use
                       'must be declared

Private Sub cboField_Click()
  If cboField.Text = "(All Fields)" Then
     chkMatch.Value = 0
     chkMatch.Enabled = False
  Else
     chkMatch.Enabled = True
  End If
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cboFilter_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

'If there is a change in cboFilter...
Private Sub cboFilter_Change()
  If Len(Trim(cboFilter.Text)) > 0 Then
     'cmdFilter will be active and ready
     cmdFilter.Enabled = True
     cmdFilter.Default = True
  Else 'Still empty
     cmdFilter.Enabled = False 'We can't use it
  End If
End Sub

Private Sub cmdFilter_Click()
On Error GoTo Message
 'Assign recordset variable to new recordset
  Set adoFilter = New ADODB.Recordset
 'Filter recordset based on paramter in SQL Statement
 AddCriteriaToCombo
 If cboField.Text <> "(All Fields)" Then
   If chkMatch.Value = 0 Then 'Not match whole criteria word

     adoFilter.Open "SHAPE " & _
     "{SELECT * FROM " & m_RecordSource1 & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "LIKE '%" & cboFilter.Text & "%' " & _
     "ORDER BY " & m_FieldKey1 & "} AS ParentCMD APPEND " & _
     "({SELECT * FROM " & m_RecordSource1 & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "LIKE '%" & cboFilter.Text & "%' " & _
     "ORDER BY " & m_FieldKey1 & "} AS ChildCMD " & _
     "RELATE " & m_FieldKey1 & " TO " & m_FieldKey1 & ") " & _
     "AS ChildCMD", cnn, adOpenStatic, adLockOptimistic

   Else 'Match whole criteria word only
     adoFilter.Open "SHAPE " & _
     "{SELECT * FROM " & m_RecordSource1 & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "= '" & cboFilter.Text & "' " & _
     "ORDER BY " & m_FieldKey1 & "} AS ParentCMD APPEND " & _
     "({SELECT * FROM " & m_RecordSource1 & " " & _
     "WHERE " & Trim(cboField.Text) & " " & _
     "= '" & cboFilter.Text & "' " & _
     "ORDER BY " & m_FieldKey1 & "} AS ChildCMD " & _
     "RELATE " & m_FieldKey1 & " TO " & m_FieldKey1 & ") " & _
     "AS ChildCMD", cnn, adOpenStatic, adLockOptimistic
   End If

   With frmPersonal
     'If recordset is not empty
     If adoFilter.RecordCount > 0 Then
       'Display the result to datagrid

       'This will update the status label in
       'middle of navigation button
       Set .adoPrimaryRS = adoFilter
       Set .grdDataGrid.DataSource = adoFilter.DataSource

       'Bind the data to textbox
       Dim oTextData As TextBox
       For Each oTextData In .txtFields
           Set oTextData.DataSource = adoFilter.DataSource
       Next
       'Go to the first record
       .cmdFirst.Value = True
       .cmdBookmark.Enabled = False
       Set .adoPrimaryRS = adoFilter
     Else
       .cmdRefresh.Value = True
       MsgBox "'" & cboFilter.Text & "' not found " & _
              "in field " & cboField.Text & ".", _
              vbExclamation, "No Result"
     End If
   End With

   Exit Sub
 Else
   FilterInAllFields
   Exit Sub
 End If
Message:
  MsgBox "'" & cboFilter.Text & "' not found " & _
         "in field '" & cboField.Text & "'.", _
         vbExclamation, "No Result"
End Sub
Private Sub cmdCancel_Click()
  'Clear memory from object variable
  Set adoField1 = Nothing
  Set rs1 = Nothing
  Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
  If cboFilter.Text = "" Then
     cmdFilter.Enabled = False
  Else 'If cboFilter is not empty
     cmdFilter.Enabled = True 'cmdFilter ready!
  End If

  Set rs1 = New ADODB.Recordset
  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic
  cboField.Clear
  cboField.AddItem "(All Fields)"
  For Each adoField1 In rs1.Fields
      cboField.AddItem adoField1.Name
  Next
  cboField.Text = cboField.List(0)
  'Get setting for this form from INI File
  Call ReadFromINIToControls(frmFilter, "Filter")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Save setting this form to INI File
  Call SaveFromControlsToINI(frmFilter, "Filter")
  'Clear memory
  Set adoFilter = Nothing
  Set adoField1 = Nothing
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilterInAllFields()
Dim strCriteria As String, strField As String
Dim intField As Integer, i As Integer, j As Integer
Dim tabField() As String
  rs1.MoveFirst
  strCriteria = " "
  intField = rs1.Fields.Count
  ReDim tabField(intField)
  intField = rs1.Fields.Count
  i = 0
  For Each adoField1 In rs1.Fields
      tabField(i) = adoField1.Name
      i = i + 1
  Next
  For i = 0 To intField - 1
    If chkMatch.Value = 0 Then 'Not match whole criteria word
     If i <> intField - 1 Then
        strField = strField & tabField(i) & ","
        strCriteria = strCriteria & _
           tabField(i) & " LIKE '%" & cboFilter.Text & "%' Or "
     Else

        strField = strField & tabField(i) & ""
        strCriteria = strCriteria & tabField(i) & " LIKE '%" & cboFilter.Text & "%' "
     End If
  Else  'Match whole criteria word only
     If i <> intField - 1 Then
        strField = strField & tabField(i) & ","
        strCriteria = strCriteria & _
           tabField(i) & " = '%" & cboFilter.Text & "%' Or "
     Else
        strField = strField & tabField(i) & ""
        strCriteria = strCriteria & tabField(i) & " = '%" & cboFilter.Text & "%' "
     End If
  End If
  Next i
  Set adoFilter = New ADODB.Recordset
     adoFilter.Open _
     "SHAPE " & _
     "{SELECT " & strField & " FROM " & m_RecordSource1 & " " & _
     "WHERE " & strCriteria & " ORDER BY " & m_FieldKey1 & "} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT " & strField & " FROM " & m_RecordSource1 & " " & _
     "WHERE " & strCriteria & " ORDER BY " & m_FieldKey1 & "} " & _
     "AS ChildCMD RELATE " & m_FieldKey1 & " TO " & m_FieldKey1 & ") " & _
     "AS ChildCMD", cnn, adOpenStatic, adLockOptimistic

  With frmPersonal
  If adoFilter.RecordCount > 0 Then
     Set .adoPrimaryRS = adoFilter
     Set .grdDataGrid.DataSource = adoFilter.DataSource
     Dim oTextData As TextBox
     For Each oTextData In .txtFields
         Set oTextData.DataSource = adoFilter.DataSource
     Next
     .cmdFirst.Value = True
     .cmdBookmark.Enabled = False
  Else
     .cmdRefresh.Value = True
     MsgBox "'" & cboFilter.Text & "' not found " & _
            "in field '" & cboField.Text & "'.", _
            vbExclamation, "No Result"
  End If
  End With
  Exit Sub
Message:
  'MsgBox Err.Number & " - " & Err.Description
  MsgBox "'" & cboFilter.Text & "' not found " & vbCrLf & _
         "in field '" & cboField.Text & "'.", _
         vbExclamation, "No Result"
End Sub

Private Sub AddCriteriaToCombo()
Dim i As Integer
  If cboFilter.Text = "" Then
     MsgBox "Data is empty!", _
            vbExclamation, "Empty"
     cboFilter.SetFocus
     Exit Sub
  End If
  For i = 0 To cboFilter.ListCount - 1
    If cboFilter.List(i) = cboFilter.Text Then
       cboFilter.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
  Next i
  cboFilter.AddItem cboFilter.Text
  cboFilter.Text = cboFilter.List(cboFilter.ListCount - 1)
End Sub

