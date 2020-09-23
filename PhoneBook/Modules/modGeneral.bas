Attribute VB_Name = "modGeneral"



Option Explicit

Public cnn As ADODB.Connection
Public adoFind As ADODB.Recordset
Public adoFilter As ADODB.Recordset
Public adoSort As ADODB.Recordset
Public m_ConnectionString As String
Public m_RecordSource1 As String
Public m_SQLRS1 As String
Public m_FieldKey1 As String
Public strSQL As String
Public intMax As Integer
Public rs1 As ADODB.Recordset
Public adoField1 As ADODB.Field

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Public INIFileName As String

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub OpenConnection()
  Set cnn = New ADODB.Connection
  cnn.CursorLocation = adUseClient
  m_ConnectionString = "PROVIDER=MSDataShape;Data PROVIDER=" & _
           "Microsoft.Jet.OLEDB.4.0;Data Source=" & _
           App.Path & "\DataBase\DataBase.mdb;Jet OLEDB:" & _
           "Database Password=;"
  cnn.Open m_ConnectionString
  m_SQLRS1 = "SELECT ID,FirstName,Middle,Last,Title,NickName,Email,Comments,StreetAddress,City,State,ZipCode,Country,Phone,Fax,Mobile,Web FROM Personal"
  m_RecordSource1 = "Personal"
  m_FieldKey1 = "FirstName"

End Sub

Public Sub AdjustDataGridColumnWidth _
           (DG As DataGrid, _
           adoData As ADODB.Recordset, _
           intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)

'This procedure will adjust DataGrids column width
'based on longest field in underlying source


    Dim row As Long, col As Long
    Dim Width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String
    Dim i As Integer
    'If number of records = 0 then exit from the sub
    If intRecord = 0 Then Exit Sub
    'Save the form's font for DataGrid's font
    'We need this for form's TextWidth method
    Set saveFont = DG.Parent.font
    Set DG.Parent.font = DG.font
    'Adjust ScaleMode to vbTwips for the form (parent).
    saveScaleMode = DG.Parent.ScaleMode
    DG.Parent.ScaleMode = vbTwips
    'Always from first record...
    adoData.MoveFirst
    maxWidth = 0

    'Get maximal value for progressbar control
    intMax = intField * intRecord
    frmPersonal.prgBar1.Visible = True
    frmPersonal.prgBar1.Max = intMax
        
    'We begin from the first column until the last column
    For col = 0 To intField - 1
        'Tampilkan nama field/kolom yg sedang diproses
        frmPersonal.lblField.Caption = _
           "Column: " & DG.Columns(col).DataField & ""
        adoData.MoveFirst
        'Optional param, if true, set maxWidth to
        'width of DG.Parent
        If AccForHeaders Then
            maxWidth = DG.Parent.TextWidth(DG.Columns(col).Text) + 200
        End If
        'Repeat from first record again after we have
        'finished process the last record in
        'former column...
        adoData.MoveFirst
        For row = 0 To intRecord - 1
            'Get the text from the DataGrid's cell
            If intField = 1 Then
            Else  'If number of field more than one

               cellText = DG.Columns(col).Text
            End If
            Width = DG.Parent.TextWidth(cellText) + 200
            If Width > maxWidth Then
               maxWidth = Width
               DG.Columns(col).Width = maxWidth
            End If
            adoData.MoveNext
            DoEvents
            i = i + 1
            frmPersonal.lblAngka.Caption = _
              "Finished " & Format((i / intMax) * 100, "0") & "%"
             DoEvents
            frmPersonal.prgBar1.Value = i
            DoEvents
        Next row
        DG.Columns(col).Width = maxWidth
    Next col
    'Change the DataGrid's parent property
    Set DG.Parent.font = saveFont
    DG.Parent.ScaleMode = saveScaleMode
    adoData.MoveFirst
    ResetProgressBar
End Sub  'End of AdjustDataGridColumnWidth

Public Sub ResetProgressBar()
  With frmPersonal
    .prgBar1.Value = 0
    .lblAngka.Caption = ""
    .lblField.Caption = ""
  End With
End Sub

Public Function SaveFromControlsToINI(Objek, MyAppName As String)
Dim Contrl As Control, Result As Long
Dim TempControlName As String, TempControlValue As String
On Error Resume Next
For Each Contrl In Objek
  If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Then
    TempControlName = Contrl.Name
    TempControlValue = Contrl.Value
    If (TypeOf Contrl Is ComboBox) Then
      TempControlValue = Contrl.Text
      If TempControlValue = "" Then TempControlValue = 1
    End If
    Result = WritePrivateProfileString(MyAppName, TempControlName, _
    TempControlValue, INIFileName)
  End If

  If (TypeOf Contrl Is TextBox) Then
    TempControlName = Contrl.Name
    TempControlValue = Contrl.Text
    Result = WritePrivateProfileString(MyAppName, TempControlName, _
    TempControlValue, INIFileName)
  End If
  If (TypeOf Contrl Is OptionButton) Then
    TempControlValue = Contrl.Value
    If TempControlValue = True Then
      TempControlName = Contrl.Name
      TempControlValue = Contrl.Index
      Result = WritePrivateProfileString(MyAppName, TempControlName, _
      TempControlValue, INIFileName)
    End If
  End If
Next
End Function

Public Function ReadFromINIToControls(Objek, MyAppName As String)
Dim Contrl As Control, Result As Long
Dim TempControlName As String * 101, TempControlValue As String * 101
On Error Resume Next
For Each Contrl In Objek
If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Or (TypeOf _
Contrl Is OptionButton) Or (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is CheckBox) Then
TempControlName = Contrl.Name
If (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is ComboBox) Then 'Or _
   '(TypeOf Contrl Is MaskEdBox) Then
   Result = GetPrivateProfileString(MyAppName, TempControlName, "", _
   TempControlValue, Len(TempControlValue), INIFileName)
Else 'If (TypeOf Contrl Is CheckBox) Then
   Result = GetPrivateProfileString(MyAppName, TempControlName, "0", _
   TempControlValue, Len(TempControlValue), INIFileName)
End If

If (TypeOf Contrl Is OptionButton) Then
   If Contrl.Index = Val(TempControlValue) Then Contrl = True
Else
    Contrl = TempControlValue
   If (TypeOf Contrl Is ComboBox) Then
      If Len(Contrl.Text) = 0 Then Contrl.ListIndex = 0
      End If
   End If
End If
Next
End Function

