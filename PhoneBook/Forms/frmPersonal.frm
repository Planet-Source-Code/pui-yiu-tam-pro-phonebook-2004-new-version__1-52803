VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Data - A PhoneBook 2004 - www.xoftwares.com"
   ClientHeight    =   9405
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmPersonal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   8175
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   7935
      Begin VB.PictureBox picStatBox 
         Height          =   600
         Left            =   120
         ScaleHeight     =   540
         ScaleWidth      =   5955
         TabIndex        =   26
         Top             =   6960
         Width           =   6015
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
            Height          =   350
            Left            =   5160
            TabIndex        =   30
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next"
            Height          =   350
            Left            =   4440
            TabIndex        =   29
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Prev"
            Height          =   350
            Left            =   840
            TabIndex        =   28
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
            Height          =   350
            Left            =   120
            TabIndex        =   27
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   120
            Width           =   3120
         End
      End
      Begin VB.PictureBox picButtons 
         Height          =   5985
         Left            =   6360
         ScaleHeight     =   5925
         ScaleWidth      =   1335
         TabIndex        =   23
         Top             =   240
         Width           =   1395
         Begin PhoneBook.chameleonButton cmdPrint 
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   5400
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Print"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":08CA
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdClose 
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   4920
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Close"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":08F6
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdDataGrid 
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   4440
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Grid"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":0922
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdSort 
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   3960
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Sort"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":094E
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdFilter 
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   3480
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Filter"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":097A
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdFind 
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   3000
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Find"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":09A6
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdRefresh 
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   2520
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Refresh"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":09D2
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdDelete 
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   2040
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Delete"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":09FE
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdEdit 
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1560
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Edit"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":0A2A
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdCancel 
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1080
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Cancel"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":0A56
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdUpdate 
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   600
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Update"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":0A82
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
         Begin PhoneBook.chameleonButton cmdAdd 
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   1095
            _extentx        =   1931
            _extenty        =   661
            btype           =   3
            tx              =   "&Add"
            enab            =   -1  'True
            font            =   "frmPersonal.frx":0AAE
            coltype         =   1
            focusr          =   -1  'True
            bcol            =   14215660
            fcol            =   0
         End
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Web"
         Height          =   285
         Index           =   16
         Left            =   1920
         TabIndex        =   22
         Top             =   5235
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Mobile"
         Height          =   285
         Index           =   15
         Left            =   1920
         TabIndex        =   21
         Top             =   4920
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Fax"
         Height          =   285
         Index           =   14
         Left            =   1920
         TabIndex        =   20
         Top             =   4605
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Phone"
         Height          =   285
         Index           =   13
         Left            =   1920
         TabIndex        =   19
         Top             =   4290
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Country"
         Height          =   285
         Index           =   12
         Left            =   1920
         TabIndex        =   18
         Top             =   3990
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ZipCode"
         Height          =   285
         Index           =   11
         Left            =   1920
         TabIndex        =   17
         Top             =   3675
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "State"
         Height          =   285
         Index           =   10
         Left            =   1920
         TabIndex        =   16
         Top             =   3360
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "City"
         Height          =   285
         Index           =   9
         Left            =   1920
         TabIndex        =   15
         Top             =   3045
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "StreetAddress"
         Height          =   285
         Index           =   8
         Left            =   1920
         TabIndex        =   14
         Top             =   2730
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Comments"
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   13
         Top             =   2430
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Email"
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   12
         Top             =   2115
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "NickName"
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   11
         Top             =   1800
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Title"
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   10
         Top             =   1485
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Last"
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   1170
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Middle"
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   8
         Top             =   870
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "FirstName"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   7
         Top             =   555
         Width           =   4230
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ID"
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   4230
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   780
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5595
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1376
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ProgressBar prgBar1 
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   6720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   318
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   6480
         Width           =   2655
      End
      Begin VB.Label lblAngka 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4680
         TabIndex        =   49
         Top             =   7650
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site:"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   5235
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   47
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   46
         Top             =   4605
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   45
         Top             =   4290
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   44
         Top             =   3990
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   43
         Top             =   3675
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   41
         Top             =   3045
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Street Address:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   40
         Top             =   2730
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   39
         Top             =   2430
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   2115
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Name:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   1485
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   555
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "ID:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9135
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8943
            MinWidth        =   7408
            Object.Tag             =   ""
            Object.ToolTipText     =   "(C) Masino Sinaga (masino_sinaga@yahoo.com)"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Object.Tag             =   ""
            Object.ToolTipText     =   "It's up to you..."
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "4/1/2004"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Date today"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1464
            MinWidth        =   1464
            TextSave        =   "11:57 PM"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Time right now"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -720
      X2              =   8160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use the below options to work with A PhoneBook 2004"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   3915
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPersonal.frx":0ADA
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1.00005e5
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Data - PhoneBook 2004"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   4155
   End
   Begin VB.Image imgLogo 
      Height          =   1335
      Left            =   5160
      Picture         =   "frmPersonal.frx":0F1C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3330
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'General variable for this module
Public WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Public WithEvents rsstrFindData As Recordset
Attribute rsstrFindData.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim blnCancel As Boolean
Dim NumData As Integer
Dim intRecord As Integer
Dim intField As Integer

Private Sub cmdPrint_Click()
    Load rptList
    rptList.Show
End Sub

Private Sub Form_Load()
On Error GoTo Message
INIFileName = App.Path & "\SettingfrmPersonal.ini"
  blnCancel = False
  OpenConnection
  Set adoPrimaryRS = New Recordset
  'We display all data in a datagrid below and underlying
  'source (the selected record in datagrid) above.
  strSQL = "SHAPE {SELECT ID,FirstName,Middle,Last,Title,NickName,Email,Comments,StreetAddress,City,State,ZipCode,Country,Phone,Fax,Mobile,Web " & vbCrLf & _
  "FROM Personal   ORDER BY FirstName } AS ParentCMD APPEND " & vbCrLf & _
  "({SELECT ID,FirstName,Middle,Last,Title,NickName,Email,Comments,StreetAddress,City,State,ZipCode,Country,Phone,Fax,Mobile,Web " & vbCrLf & _
  "FROM Personal  } AS ChildCMD RELATE FirstName  TO FirstName) AS ChildCMD"
  adoPrimaryRS.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
  Dim oText As TextBox
  'Bind textbox to recordset
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  'Bind recordset to datagrid
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  mbDataChanged = False

  LockTheForm  'Lock textbox, and make datagrid enable
  grdDataGrid.Enabled = True
  'If we have no data in recordset
  If adoPrimaryRS.RecordCount < 1 Then
     MsgBox "Recordset is empty. Please click Add button to add new record!", vbExclamation, "Empty Recordset"
     Exit Sub
  End If
  LockTheForm 'Lock textbox, combobox, and optionbutton
  'Except Datagrid....
  grdDataGrid.Enabled = True
  grdDataGrid.TabStop = False
  SetButtons True
  Exit Sub
Message:
  MsgBox Err.Number & " - " & Err.Description
  End
End Sub

Private Sub Message(strMessage As String)
  StatusBar1.Panels(1).Text = strMessage
End Sub


Private Sub cmdDataGrid_Click()
  intRecord = adoPrimaryRS.RecordCount
  intField = adoPrimaryRS.Fields.Count - 1
  Call AdjustDataGridColumnWidth(grdDataGrid, adoPrimaryRS, _
                              intRecord, intField, True)
End Sub
Private Sub cmdDataGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Adjust datagrid columns based on the longest field.")
End Sub


Private Sub cmdFirst_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Go to the first record.")
End Sub

Private Sub cmdLast_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Go to the last record.")
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Go to the next record.")
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Go to the previous record.")
End Sub


Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
  Response = -1
  'DataError = -1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then
     MsgBox "You have to save or cancel the changes " & vbCrLf & _
            "that you have just made before quit!", _
            vbExclamation, "Warning"
     cmdUpdate.SetFocus
     Cancel = -1
     Exit Sub
  End If

  If Not adoPrimaryRS Is Nothing Then _
    Set adoPrimaryRS = Nothing  'Clear memory from recordset
  'In order that prevent error from DataGrid...!
  If grdDataGrid.TabStop = True Then
     txtFields(0).SetFocus
  End If
  cnn.Close 'Close database
  Set cnn = Nothing  'Clear memory from database
  End
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault 'Mouse pointer back to normal
End Sub

'Display the selected record in datagrid
Public Sub adoPrimaryRS_MoveComplete(ByVal adReason As _
            ADODB.EventReasonEnum, ByVal pError As _
            ADODB.Error, adStatus As ADODB.EventStatusEnum, _
            ByVal pRecordset As ADODB.Recordset)
  NumData = adoPrimaryRS.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(NumData) & " from " _
                      & adoPrimaryRS.RecordCount
  CheckNavigation
End Sub

Private Sub CheckNavigation()
  'This will check which navigation button can be
  'accessed when you navigate the recordset through
  'Datagrid control or navigation button itself
  With adoPrimaryRS
   'If we have at least two record...
   If (.RecordCount > 1) Then
      'BOF = Begin Of Recordset
      If (.BOF) Or _
         (.AbsolutePosition = 1) Then
          cmdFirst.Enabled = False
          cmdPrevious.Enabled = False
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      'EOF = End Of Recordset
      ElseIf (.EOF) Or _
          (.AbsolutePosition = .RecordCount) Then
          cmdNext.Enabled = False
          cmdLast.Enabled = False
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True

      Else
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      End If
   Else
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
 End With
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    UnlockTheForm
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
  grdDataGrid.Enabled = False  'In order that prevent error
  On Error Resume Next
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Add new record.")
End Sub


Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  If adoPrimaryRS.RecordCount < 1 Then
     MsgBox "Recordset is empty. Please click Add button to add new record!", vbExclamation, "Empty Recordset"
     Exit Sub
  End If
  If MsgBox("Are you sure you want to delete this record?", _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            "Delete Record") _
            <> vbYes Then
     Exit Sub
  End If
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Delete the selected record.")
End Sub


Private Sub cmdRefresh_Click()
  'Refresh is very important in multiuser app
  On Error GoTo RefreshErr
  If blnCancel = True Then
     SetButtons True
     blnCancel = False
  End If
  LockTheForm
  Set grdDataGrid.DataSource = Nothing
  Set adoPrimaryRS = New Recordset
  strSQL = "SHAPE {SELECT ID,FirstName,Middle,Last,Title,NickName,Email,Comments,StreetAddress,City,State,ZipCode,Country,Phone,Fax,Mobile,Web " & vbCrLf & _
  "FROM Personal   ORDER BY FirstName } AS ParentCMD APPEND " & vbCrLf & _
  "({SELECT ID,FirstName,Middle,Last,Title,NickName,Email,Comments,StreetAddress,City,State,ZipCode,Country,Phone,Fax,Mobile,Web " & vbCrLf & _
  "FROM Personal  } AS ChildCMD RELATE FirstName  TO FirstName) AS ChildCMD"
  adoPrimaryRS.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  

  'Bind recordset to datagrid
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  grdDataGrid.Enabled = True
  Exit Sub

RefreshErr:
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark <> 0 Then
      adoPrimaryRS.Bookmark = mvBookMark
  Else
      adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  blnCancel = True
  cmdRefresh_Click  'Automatically refresh
  Exit Sub
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Retrieve all records from database.")
End Sub


Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  If adoPrimaryRS.RecordCount < 1 Then
     MsgBox "Recordset is empty. Please click Add button to add new record!", vbExclamation, "Empty Recordset"
     Exit Sub
  End If
  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  UnlockTheForm 'Unlock textbox; we can edit data
  txtFields(0).SetFocus: SendKeys "{Home}+{End}"
  Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Edit the selected record.")
End Sub


Private Sub cmdCancel_Click()
  On Error Resume Next
  LockTheForm
  cmdRefresh_Click
  grdDataGrid.Enabled = True
  If blnCancel = True Then
     Exit Sub
  End If
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  LockTheForm    'Lock textbox
  grdDataGrid.Enabled = True
  mbDataChanged = False
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Cancel the change or new record that have not been saved.")
End Sub


Private Sub cmdUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Save the change or new record.")
End Sub


Private Sub cmdUpdate_Click()
Dim i As Integer
  On Error GoTo UpdateErr
  For i = 0 To 16
    If txtFields(i).Text = "" Then
       MsgBox "You have to fill in all textbox!", _
              vbExclamation, "Validation"
       txtFields(i).SetFocus
       Exit Sub
     End If
  Next i
  'Update by using UpdateBatch. UpdateBatch will
  'automatically update all data in various fields type.
  adoPrimaryRS.UpdateBatch adAffectAll
  'Move pointer to last record if we just added data
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast
  End If
  If mbEditFlag Then
    adoPrimaryRS.MoveNext
    adoPrimaryRS.MovePrevious
  End If

  'Update all status
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  LockTheForm  'Lock textbox
  grdDataGrid.Enabled = True
  'Display the record position
  NumData = adoPrimaryRS.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(NumData) & " from " _
                      & adoPrimaryRS.RecordCount
  Exit Sub
UpdateErr:
  MsgBox Err.Number & " - " & _
         Err.Description, vbCritical, "Error Occured"
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Quit from this program now.")
End Sub


Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
If adoFilter Is Nothing Then
   adoPrimaryRS.MoveFirst
Else
   adoFilter.MoveFirst
End If
  mbDataChanged = False
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
If adoFilter Is Nothing Then
   adoPrimaryRS.MoveLast
Else
   adoFilter.MoveLast
End If
  mbDataChanged = False
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
If adoFilter Is Nothing Then
   If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
   If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
      Beep
      adoPrimaryRS.MoveLast
      MsgBox "This is the last record.", _
             vbInformation, "Last Record"
   End If
Else
   If Not adoFilter.EOF Then adoFilter.MoveNext
   If adoFilter.EOF And adoFilter.RecordCount > 0 Then
      Beep
      adoFilter.MoveLast
      MsgBox "This is the last record.", _
             vbInformation, "Last Record"
   End If
End If
  mbDataChanged = False
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub


Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
If adoFilter Is Nothing Then
   If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
   If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
      Beep
      adoPrimaryRS.MoveFirst
      MsgBox "This is the first record.", _
             vbInformation, "First Record"
   End If
Else
   If Not adoFilter.BOF Then adoFilter.MovePrevious
   If adoFilter.BOF And adoFilter.RecordCount > 0 Then
      Beep
      adoFilter.MoveFirst
      MsgBox "This is the first record.", _
             vbInformation, "First Record"
   End If
End If
  mbDataChanged = False
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)

  cmdAdd.Enabled = bVal
  cmdUpdate.Enabled = Not bVal
  cmdCancel.Enabled = Not bVal
  cmdEdit.Enabled = bVal
  cmdDelete.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdFind.Enabled = bVal
  cmdFilter.Enabled = bVal
  cmdSort.Enabled = bVal
  cmdDataGrid.Enabled = bVal
  cmdClose.Enabled = bVal

  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub



Private Sub picButtons_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index  'If we hit Enter, jump to next textbox
         Case 0 To 16
              If KeyAscii = 13 Then SendKeys "{Tab}"
  End Select
End Sub

'Lock textbox in order that we can't edit data
Private Sub LockTheForm()
Dim i As Integer
  For i = 0 To 16
    txtFields(i).Locked = True
  Next i
  grdDataGrid.Enabled = False
End Sub

'Unlock textbox in order that we can edit data
Sub UnlockTheForm()
Dim i As Integer
  For i = 0 To 16
    txtFields(i).Locked = False
  Next i
  grdDataGrid.Enabled = False
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  Set adoFilter = Nothing
  Set adoSort = Nothing
  Set adoFind = New ADODB.Recordset
  Set adoFind = adoPrimaryRS
  frmFind.Show , frmPersonal
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Find record (find first and find next).")
End Sub

Private Sub cmdFilter_Click()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  Set adoFind = Nothing
  Set adoSort = Nothing
  Set adoFilter = New ADODB.Recordset
  Set adoFilter = adoPrimaryRS
  frmFilter.Show , frmPersonal
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdFilter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Filter recordset.")
End Sub

Public Sub rsstrFindData_MoveComplete(ByVal adReason As _
            ADODB.EventReasonEnum, ByVal pError As _
            ADODB.Error, adStatus As ADODB.EventStatusEnum, _
            ByVal pRecordset As ADODB.Recordset)
    NumData = rsstrFindData.AbsolutePosition
    lblStatus.Caption = "Record number " & CStr(NumData) & " from " _
                      & rsstrFindData.RecordCount
End Sub

Private Sub cmdSort_Click()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  Set adoFind = Nothing
  Set adoFilter = Nothing
  Set adoSort = New ADODB.Recordset
  Set adoSort = adoPrimaryRS
  frmSort.Show , frmPersonal
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Call Message("Sort recordset.")
End Sub

