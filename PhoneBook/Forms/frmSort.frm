VERSION 5.00
Begin VB.Form frmSort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5655
      Begin PhoneBook.chameleonButton cmdCancel 
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
      Begin PhoneBook.chameleonButton cmdSort 
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Sort"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         FCOL            =   0
      End
      Begin VB.ComboBox cboField 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cboSort 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort in Field:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
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
      Picture         =   "frmSort.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Data - PhoneBook 2004"
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
      Width           =   2460
   End
   Begin VB.Image imgLogo 
      Height          =   1335
      Left            =   2760
      Picture         =   "frmSort.frx":0442
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
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboField_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
Private Sub cboSort_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
Private Sub cmdCancel_Click()
  'Clear memory from object variable
  Set adoField1 = Nothing
  Set rs1 = Nothing
  Unload Me
End Sub

Private Sub cmdSort_Click()
Dim TipeSort As String
   If cboSort.Text = cboSort.List(0) Then
      TipeSort = "ASC"
   Else
      TipeSort = "DESC"
   End If
   Set adoSort = New ADODB.Recordset
   adoSort.Open "SHAPE " & _
     "{SELECT * FROM " & m_RecordSource1 & " " & _
     "ORDER BY " & cboField.Text & " " & TipeSort & "} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT * FROM " & m_RecordSource1 & " " & _
     "ORDER BY " & cboField.Text & " " & TipeSort & "} " & _
     "AS ChildCMD RELATE " & m_FieldKey1 & " TO " & m_FieldKey1 & " ) " & _
     "AS ChildCMD", _
     cnn, adOpenStatic, adLockOptimistic
   With frmPersonal
     On Error Resume Next
     If adoSort.RecordCount > 0 Then
        Set .grdDataGrid.DataSource = adoSort.DataSource
        Set .rsstrFindData = adoSort.DataSource
        Dim oTextData As TextBox

        For Each oTextData In .txtFields
            Set oTextData.DataSource = adoSort.DataSource
        Next
        Set .adoPrimaryRS = adoSort
        .cmdFirst.Value = True
     End If
   End With
   Exit Sub
Message:
     MsgBox Err.Number & " - " & _
            Err.Description, _
            vbExclamation, "No Result"
End Sub

Private Sub Form_Load()
On Error Resume Next
  Set rs1 = New ADODB.Recordset
  rs1.Open m_SQLRS1, cnn, adOpenKeyset, adLockOptimistic
  cboField.Clear
  For Each adoField1 In rs1.Fields
      cboField.AddItem adoField1.Name
  Next
  rs1.Close
  cboField.Text = cboField.List(0)
  cboSort.AddItem "Ascending (ASC)"
  cboSort.AddItem "Descending (DESC)"
  cboSort.Text = cboSort.List(0)
  'Get setting for this form from INI File
  Call ReadFromINIToControls(frmSort, "Sort")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Save setting this form to INI File
  Call SaveFromControlsToINI(frmSort, "Sort")
  'Clear memory
  Set adoSort = Nothing
  Set adoField1 = Nothing
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

