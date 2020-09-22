VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar Bar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   4635
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   5520
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ProxyChecker.IPTextBox txtIP 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   661
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   6240
      Top             =   3240
   End
   Begin MSComctlLib.ListView lsv1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IP"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Location"
         Object.Width           =   4586
      EndProperty
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FileNameSt As String

Public Function Run(StrFileName As String)
    
    FileNameSt = StrFileName
    Me.Show vbModal

End Function


Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdSave_Click()

    cmdSave.Enabled = False
'
    If lsv1.ListItems.Count = 0 Then Exit Sub
    
    Bar1.Max = lsv1.ListItems.Count
    
    With Adodc1
        .ConnectionString = dbPath
        For i = 1 To lsv1.ListItems.Count
            If lsv1.ListItems(i).Checked = True Then
                .RecordSource = "select * from proxy where ( ip = '" & _
                lsv1.ListItems(i).Text & "' ) and ( port = '" & lsv1.ListItems(i).SubItems(1) & _
                "') "
                .Refresh
                If .Recordset.RecordCount = 0 Then
'                    .Recordset.ActiveConnection = Nothing
'                    .RecordSource = "select * from proxy"
                    .Refresh
                    .Recordset.AddNew
                    .Recordset.Fields(1) = lsv1.ListItems(i).Text
                    .Recordset.Fields(2) = lsv1.ListItems(i).SubItems(1)
                    .Recordset.Fields(3) = lsv1.ListItems(i).SubItems(2)
                    .Recordset.Fields(4) = Now
                    .Recordset.Update
                End If
                .Recordset.ActiveConnection = Nothing
                
            End If
            Bar1.Value = Bar1.Value + 1
        Next i
    End With
    
    Bar1.Value = 0
    
    MsgBox LoadResString(IsFar * 1000 + 117), vbOKOnly, ""
    
    frm1.TimerList.Enabled = True
    
    Unload Me
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 And cmdSave.Enabled = True Then cmdSave_Click
    If KeyCode = 27 Then Exit Sub

End Sub


Private Sub Form_Load()

For i = 0 To 2
    lsv1.ColumnHeaders(i + 1).Text = LoadResString(IsFar * 1000 + i + 148)
Next i

With Me
    .Caption = LoadResString(IsFar * 1000 + 123)
    .RightToLeft = CBool(IsFar)
End With

cmdSave.Caption = LoadResString(IsFar * 1000 + 122)
cmdExit.Caption = LoadResString(IsFar * 1000 + 164)



End Sub

Private Sub Timer1_Timer()

    Dim st As String
    Dim IP As String
    Dim PORT As String
    Dim LOCATION As String
    Dim i As Integer
    
    i = 0
    
    Timer1.Enabled = False

    Open FileNameSt For Input As #1
        Do While Not EOF(1)
            Input #1, st
            If InStr(1, st, ":", vbTextCompare) <> 0 Then
                IP = Split(st, ":")(0)
                txtIP.IP = IP
                If txtIP.IP = IP Then
                    st = Split(st, ":")(1)
                    PORT = Split(st, vbTab)(0)
                    If IsNumeric(PORT) = True Then
                        st = Split(st, vbTab)(1)
                        If st = "" Then st = "UnKnown"
                        LOCATION = st
                        i = i + 1
                        With lsv1
                            .ListItems.Add i, , IP, , 1
                            .ListItems(i).ListSubItems.Add 1, , PORT
                            .ListItems(i).ListSubItems.Add 2, , LOCATION
                            .ListItems(i).Checked = True
                        End With
                    End If
                End If
            End If
        Loop
    Close #1
    
    cmdExit.Enabled = True
    If lsv1.ListItems.Count <> 0 Then cmdSave.Enabled = True

End Sub
