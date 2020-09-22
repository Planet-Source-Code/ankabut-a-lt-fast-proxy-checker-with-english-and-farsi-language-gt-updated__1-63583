VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit  Proxy"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   3240
      Top             =   1800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   1920
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
      Caption         =   "Exit"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Fra1 
      Caption         =   "Location :"
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3615
      Begin VB.TextBox txtLocation 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   90
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Fra1 
      Caption         =   "Port :"
      Height          =   735
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Fra1 
      Caption         =   "IP Address :"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      Begin ProxyChecker.IPTextBox txtIP 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   661
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim key As Integer


Public Function RunEdit(keyEdit As Integer)

    key = keyEdit
    Me.Show

End Function



Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

    If txtPort.Text = "" Then
        txtPort.SetFocus
        Beep
        Exit Sub
    End If
    
    With Adodc1
        .ConnectionString = dbPath
        If txtLocation.Text = "" Then txtLocation.Text = "UnKnown"
        .RecordSource = "select  * from proxy where key = " & key
        .Refresh
        .Recordset.AddNew
        .Recordset.Fields(1) = txtIP.IP
        .Recordset.Fields(2) = txtPort.Text
        .Recordset.Fields(3) = txtLocation.Text
        .Recordset.Update
        .Recordset.ActiveConnection = Nothing
    End With
        
        frm1.TimerList.Enabled = True
        Unload Me
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then cmdExit_Click
    If KeyCode = 13 Then cmdSave_Click
    
End Sub



Private Sub Form_Load()
With Me
    .Caption = LoadResString(IsFar * 1000 + 165)
    .RightToLeft = CBool(IsFar)
End With

For i = 0 To 2
    Fra1(i).RightToLeft = CBool(IsFar)
    Fra1(i).Caption = LoadResString(1000 * IsFar + 148 + i)
Next i

txtIP.RightToLeft = CBool(IsFar)
txtPort.RightToLeft = CBool(IsFar)

cmdSave.RightToLeft = CBool(IsFar)
cmdSave.Caption = LoadResString(IsFar * 1000 + 142)
cmdExit.RightToLeft = CBool(IsFar)
cmdExit.Caption = LoadResString(IsFar * 1000 + 125)

End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False

    With Adodc1
        .ConnectionString = dbPath
        .RecordSource = "select * from proxy where key = " & key
        .Refresh
        txtIP.IP = .Recordset.Fields(1)
        txtPort.Text = .Recordset.Fields(2)
        txtLocation.Text = .Recordset.Fields(3)
        .Recordset.ActiveConnection = Nothing
    End With

End Sub

Private Sub txtLocation_Change()
txtLocation.Text = TxtFilterToStr(txtLocation.Text)
End Sub

Private Sub txtPort_Change()
txtPort.Text = TxtFilterToNumber(txtPort.Text)
End Sub
