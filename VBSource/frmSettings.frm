VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra2 
      Caption         =   "IE Proxy :"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox chk1 
         Caption         =   "Enable Proxy Connection ."
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2280
         MaxLength       =   7
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin ProxyChecker.IPTextBox txtIP 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   661
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame fra1 
      Caption         =   "Time Out :"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
      Begin VB.ComboBox cmbTime 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk1_Click()

    If chk1.Value = Checked Then
            txtIP.Enabled = True
            txtPort.Enabled = True
        Else
            txtIP.Enabled = False
            txtPort.Enabled = False
    End If
    

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

    Dim ValueEnable As Integer


'

    If chk1.Value = Checked And (txtIP.IP = "" Or txtPort = "") Then
        Beep
        txtIP.SetFocus
        Exit Sub
    End If
    
    If chk1.Value = Checked Then ValueEnable = 1 Else ValueEnable = 0
    
    SaveProxySettings txtIP.IP, txtPort.Text, ValueEnable

    SaveSetting "ProxyChecker", "Settings", "TimeOut", Split(cmbTime.Text, "   ")(0)
    
    MsgBox LoadResString(IsFar * 1000 + 173), vbOKOnly, ""
        
    Unload Me


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then cmdSave_Click
    
    If KeyCode = 27 Then cmdExit_Click
    
End Sub

Private Sub Form_Load()


    With Me
        .Caption = LoadResString(IsFar * 1000 + 144)
        .RightToLeft = CBool(IsFar)
    End With
    
    fra1.RightToLeft = CBool(IsFar)
    fra1.Caption = LoadResString(IsFar * 1000 + 167)
    
    fra2.RightToLeft = CBool(IsFar)
    fra2.Caption = LoadResString(IsFar * 1000 + 171)
    
    chk1.RightToLeft = CBool(IsFar)
    chk1.Alignment = IsFar
    chk1.Caption = LoadResString(IsFar * 1000 + 172)
    
    cmbTime.RightToLeft = CBool(IsFar)
    
    cmdExit.Caption = LoadResString(IsFar * 1000 + 164)
    cmdSave.Caption = LoadResString(IsFar * 1000 + 122)
    
    txtIP.RightToLeft = CBool(IsFar)
    txtPort.RightToLeft = CBool(IsFar)
    
    
    For i = 1 To 45
        cmbTime.AddItem i & "   " & LoadResString(IsFar * 1000 + 118)
    Next i
    
    Dim st As String
    
    st = GetSetting("ProxyChecker", "Settings", "TimeOut")
    
    If st = "" Or Val(st) = 0 Then st = 15
    
    cmbTime.ListIndex = Val(st) - 1
    
    
 Dim Create

 'Get IE Registry Settings
 Const ProxyServer = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
 Const ProxyEnable = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"

 'Make read registry without extra modules
 Set Create = CreateObject("wscript.shell")

  'Set Old Setthings
  On Error Resume Next

  st = Create.RegRead(ProxyServer)
  txtIP.IP = Split(st, ":")(0)
  txtPort.Text = Split(st, ":")(1)
  st = Create.RegRead(ProxyEnable)
  If st = 1 Then
        chk1.Value = Checked
        txtIP.Enabled = True
        txtPort.Enabled = True
    Else
        chk1.Value = Unchecked
        txtIP.Enabled = False
        txtPort.Enabled = False
  End If

    
End Sub

Private Sub txtPort_Change()
txtPort.Text = TxtFilterToNumber(txtPort.Text)
End Sub
