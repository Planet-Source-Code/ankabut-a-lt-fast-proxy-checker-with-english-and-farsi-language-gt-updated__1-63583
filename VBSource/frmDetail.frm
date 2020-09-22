VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proxy Details"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDetail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLoc 
      Caption         =   "Location :"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   3735
      Begin VB.TextBox txtLoc 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   2640
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame fraSta 
      Caption         =   "Status :"
      Height          =   735
      Left            =   2520
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
      Begin VB.TextBox txtSta 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraLLD 
      Caption         =   "Last Live Date :"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
      Begin VB.TextBox txtLLD 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame frmDate 
      Caption         =   "Enter List Date :"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
      Begin VB.TextBox txtELD 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame frmPort 
      Caption         =   "Port :"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmIP 
      Caption         =   "IP :"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin ProxyChecker.IPTextBox txtIP 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   661
         Locked          =   -1  'True
      End
   End
End
Attribute VB_Name = "frmDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim key As Integer


Public Function RunByCode(CodStr As String)

key = Val(CodStr)

Me.Show vbModal

End Function


Private Sub cmdOK_Click()
'
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        
        Case 13
            cmdOK_Click
        Case 27
            Unload Me
    End Select

End Sub

Private Sub Form_Load()


    With Me
        .Caption = LoadResString(1000 * IsFar + 153)
        .RightToLeft = CBool(IsFar)
    End With
    
    frmIP.RightToLeft = CBool(IsFar)
    frmPort.RightToLeft = CBool(IsFar)
    fraLoc.RightToLeft = CBool(IsFar)
    fraLLD.RightToLeft = CBool(IsFar)
    fraSta.RightToLeft = CBool(IsFar)
    frmDate.RightToLeft = CBool(IsFar)
    
    txtIP.RightToLeft = CBool(IsFar)
    txtPort.RightToLeft = CBool(IsFar)
    txtSta.RightToLeft = CBool(IsFar)
    txtLLD.RightToLeft = CBool(IsFar)
    txtELD.RightToLeft = CBool(IsFar)

    frmIP.Caption = LoadResString(1000 * IsFar + 157)
    frmPort.Caption = LoadResString(1000 * IsFar + 158)
    fraLoc.Caption = LoadResString(1000 * IsFar + 159)
    fraLLD.Caption = LoadResString(1000 * IsFar + 162)
    fraSta.Caption = LoadResString(1000 * IsFar + 163)
    frmDate.Caption = LoadResString(IsFar * 1000 + 161)
    
    cmdOK.Caption = LoadResString(IsFar * 1000 + 160)
    
    With Adodc1
        .ConnectionString = dbPath
        .RecordSource = "select * from proxy where key = " & key
        .Refresh
        txtIP.IP = .Recordset.Fields(1)
        txtPort.Text = .Recordset.Fields(2)
        txtLoc.Text = .Recordset.Fields(3)
        txtELD.Text = .Recordset.Fields(4)
        If IsNull(.Recordset.Fields(5)) = False Then txtLLD.Text = .Recordset.Fields(5) Else txtLLD.Text = "-"
        If .Recordset.Fields(6) = True Then
                txtSta.Text = LoadResString(IsFar * 1000 + 101)
            Else
                txtSta.Text = LoadResString(IsFar * 1000 + 102)
        End If
        .Recordset.ActiveConnection = Nothing
    End With

End Sub
