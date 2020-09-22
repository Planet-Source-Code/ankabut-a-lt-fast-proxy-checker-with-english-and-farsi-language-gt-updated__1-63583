VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm1 
   BorderStyle     =   0  'None
   Caption         =   "Proxy Checker"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3480
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer TimerBar 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   4200
   End
   Begin MSWinsockLib.Winsock Sock1 
      Index           =   0
      Left            =   5640
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   15000
      Left            =   6000
      Top             =   3720
   End
   Begin VB.Timer TimerList 
      Interval        =   20
      Left            =   5520
      Top             =   3720
   End
   Begin MSComctlLib.ImageList img2 
      Left            =   4920
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":13FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   240
      Left            =   5970
      TabIndex        =   4
      Top             =   5120
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4320
      Top             =   4320
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
   Begin MSComctlLib.ImageList img1 
      Left            =   4320
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":1998
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":1F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":24CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":2A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":3000
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":359A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":3B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":40CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsv1 
      Height          =   4500
      Left            =   45
      TabIndex        =   3
      Top             =   495
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   7938
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Port"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Location"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Last Live Date"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Frame fra1 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   7935
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   635
      ButtonWidth     =   1614
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setting"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lang"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "English"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Farsi"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5070
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10434
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStart 
         Caption         =   "S&art Scan"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "S&top Scan"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSp101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSp102 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
      End
      Begin VB.Menu mnuSp104 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSp105 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditProxy 
         Caption         =   "&Edit Proxy"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSp107 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete Proxy"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSetting 
         Caption         =   "&Setting"
      End
      Begin VB.Menu mnuSp108 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDef 
         Caption         =   "Use as Default Proxy"
      End
      Begin VB.Menu mnuSp123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLang 
         Caption         =   "Language"
         Begin VB.Menu mnuEng 
            Caption         =   "English"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFarsi 
            Caption         =   "Farsi"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp1 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSp106 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnuDetails 
         Caption         =   "Details"
      End
      Begin VB.Menu mnusp209 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDef2 
         Caption         =   "Use as Default Proxy"
      End
      Begin VB.Menu mnuSp119 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyIP 
         Caption         =   "Copy IP"
      End
      Begin VB.Menu mnuCopyPort 
         Caption         =   "Copy Port"
      End
      Begin VB.Menu mnuSp120 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit1 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDel1 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'############################################################################
'Proxy Checker
'
'Dear friends!
'This is a free Proxy Checker . If you like my program and want me to
'continue to improve it and it's capabilities,your donations would be welcome.
'For more information please contact me !
'
'reza kargar
'
'web:    www.ragrak.com
'
'email:  kargar.reza@ gmail.com
'
'phone : +98-9122767401
'#############################################################################

Dim LastRow As Integer
Dim isRunnig As Boolean


Private Function ChangeLang()

With Status1
    .Panels(IsFar + 1).Text = ""
    .Panels(IsFar + 1).AutoSize = sbrSpring
'    .Panels(IsFar + 1).Width = .Panels(2 - IsFar).Width
    .Panels(2 - IsFar).AutoSize = sbrNoAutoSize
End With

For i = 1 To 6
    lsv1.ColumnHeaders(i).Text = LoadResString(IsFar * 1000 + 146 + i)
Next i

With Me
    .Caption = LoadResString(IsFar * 1000 + 138)
    .RightToLeft = CBool(IsFar)
End With

Bar1.Left = 5970 - IsFar * 5920

mnuFile.Caption = LoadResString(IsFar * 1000 + 119)
mnuStart.Caption = LoadResString(IsFar * 1000 + 120)
mnuStop.Caption = LoadResString(IsFar * 1000 + 121)
mnuSave.Caption = LoadResString(IsFar * 1000 + 122)
mnuImport.Caption = LoadResString(IsFar * 1000 + 123)
mnuExport.Caption = LoadResString(IsFar * 1000 + 124)
mnuExit.Caption = LoadResString(IsFar * 1000 + 125)
mnuEdit.Caption = LoadResString(IsFar * 1000 + 126)
mnuAdd.Caption = LoadResString(IsFar * 1000 + 127)
mnuEditProxy.Caption = LoadResString(IsFar * 1000 + 128)
mnuEdit1.Caption = LoadResString(IsFar * 1000 + 128)
mnuDel.Caption = LoadResString(IsFar * 1000 + 129)
mnuDel1.Caption = LoadResString(IsFar * 1000 + 129)
mnuTools.Caption = LoadResString(IsFar * 1000 + 130)
mnuSetting.Caption = LoadResString(IsFar * 1000 + 131)
mnuLang.Caption = LoadResString(IsFar * 1000 + 132)
mnuEng.Caption = LoadResString(IsFar * 1000 + 133)
mnuFarsi.Caption = LoadResString(IsFar * 1000 + 134)
mnuHelp.Caption = LoadResString(IsFar * 1000 + 135)
mnuHelp1.Caption = LoadResString(IsFar * 1000 + 136)
mnuAbout.Caption = LoadResString(IsFar * 1000 + 137)
mnuDetails.Caption = LoadResString(IsFar * 1000 + 153)
mnuCopyIP.Caption = LoadResString(IsFar * 1000 + 154)
mnuCopyPort.Caption = LoadResString(IsFar * 1000 + 155)
mnuDef.Caption = LoadResString(IsFar * 1000 + 170)
mnuDef2.Caption = LoadResString(IsFar * 1000 + 170)


With Toolbar1
    .Buttons(1).Caption = LoadResString(IsFar * 1000 + 139)
    .Buttons(2).Caption = LoadResString(IsFar * 1000 + 140)
    .Buttons(4).Caption = LoadResString(IsFar * 1000 + 141)
    .Buttons(5).Caption = LoadResString(IsFar * 1000 + 142)
    .Buttons(6).Caption = LoadResString(IsFar * 1000 + 143)
    .Buttons(8).Caption = LoadResString(IsFar * 1000 + 144)
    .Buttons(9).Caption = LoadResString(IsFar * 1000 + 145)
    .Buttons(9).ButtonMenus(1).Text = LoadResString(IsFar * 1000 + 133)
    .Buttons(9).ButtonMenus(2).Text = LoadResString(IsFar * 1000 + 134)
    .Buttons(11).Caption = LoadResString(IsFar * 1000 + 146)
End With
    
If IsFar = 0 Then
        mnuEng.Checked = True
        mnuFarsi.Checked = False
    Else
        mnuEng.Checked = False
        mnuFarsi.Checked = True
End If

End Function

Private Function CompleteScan()
    
    For i = 1 To 100
        If Timer1(i).Enabled = True Then Exit Function
    Next i
    
    IsScannig = False
    TimerBar.Enabled = False
    Bar1.Value = 0
    Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 108)
    mnuStop.Enabled = False
    mnuStart.Enabled = True
    
    

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        
        Case 13
            mnuDetails_Click
        Case 27
            mnuExit_Click
    End Select
    

End Sub

Private Sub Form_Load()

On Error Resume Next

IsFar = Val(GetSetting("ProxyChecker", "Settings", "LANG"))

ChangeLang

dbPath = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & App.Path & "\list.mdb" & ";Uid=Admin;Pwd=" & 123456 & ";"


End Sub




Private Sub Form_Unload(Cancel As Integer)

    If IsChanged = True Then
        Cancel = True
        Dim st As String
        st = LoadResString(IsFar * 1000 + 109)
        st = st & vbCrLf & vbCrLf
        st = st & LoadResString(IsFar * 1000 + 110)
        lglg = MsgBox(st, vbYesNoCancel, "")
        
        Select Case lglg
            
            Case vbYes
                Cancel = False
                mnuSave_Click
                Unload Me
            Case vbNo
                Cancel = False
                IsChanged = False
                Unload Me
            Case vbCancel
                
        End Select
    End If

End Sub


Private Sub lsv1_DblClick()
    
    frmDetail.RunByCode lsv1.SelectedItem.Text

End Sub


Private Sub lsv1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuPop
    End If
    

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuAdd_Click()
frmadd.Show vbModal
End Sub

Private Sub mnuCopyIP_Click()

    Clipboard.Clear
    Clipboard.SetText lsv1.SelectedItem.SubItems(1)

End Sub

Private Sub mnuCopyPort_Click()

    Clipboard.Clear
    Clipboard.SetText lsv1.SelectedItem.SubItems(2)

End Sub

Private Sub mnuDef_Click()
'

Dim st As String

st = LoadResString(1000 * IsFar + 174) & " "
st = st & lsv1.SelectedItem.SubItems(1) & " "
st = st & LoadResString(1000 * IsFar + 175)

lglg = MsgBox(st, vbYesNo, "")

If lglg = vbYes Then
    Dim IPStr As String
    Dim PortStr As String
    
    IPStr = lsv1.SelectedItem.SubItems(1)
    PortStr = lsv1.SelectedItem.SubItems(2)
    
    SaveProxySettings IPStr, PortStr, 1
    
    MsgBox LoadResString(IsFar * 1000 + 173), vbOKOnly, ""
End If

End Sub

Private Sub mnuDef2_Click()
mnuDef_Click
End Sub

Private Sub mnuDel_Click()

    lglg = MsgBox(LoadResString(IsFar * 1000 + 106) & lsv1.SelectedItem.Text, vbYesNo, "")
    
    If lglg = vbNo Then Exit Sub
    
    Dim con As New ADODB.Connection
    
    With con
        .ConnectionString = dbPath
        .Open
        .Execute "Delete * from proxy where key = " & Val(lsv1.SelectedItem.Text)
        .Close
    End With
    
    TimerList.Enabled = True

End Sub

Private Sub mnuDel1_Click()
mnuDel_Click
End Sub

Private Sub mnuDetails_Click()

    frmDetail.RunByCode lsv1.SelectedItem.Text

End Sub

Private Sub mnuEdit_Click()
    Status1.Panels(IsFar * 1 + 1).Text = Split(mnuEdit.Caption, "&")(1)
End Sub

Private Sub mnuEdit1_Click()
mnuEditProxy_Click
End Sub

Private Sub mnuEditProxy_Click()

    frmEdit.RunEdit Val(lsv1.SelectedItem.Text)

End Sub

Private Sub mnuEng_Click()

    If mnuEng.Checked = False Then
        IsFar = 0
        mnuEng.Checked = True
        mnuFarsi.Checked = False
        SaveSetting "ProxyChecker", "Settings", "LANG", IsFar
        ChangeLang
    End If
    

End Sub

Private Sub mnuExit_Click()

    If IsScannig = True Then
        mnuStop_Click
    End If
    
    Unload Me

End Sub

Private Sub mnuExport_Click()

    If IsChanged = True Then
        Dim st As String
        st = LoadResString(IsFar * 1000 + 109)
        st = st & vbCrLf & vbCrLf
        st = st & LoadResString(IsFar * 1000 + 113)
        MsgBox st, vbOKOnly, ""
        Exit Sub
    End If
    
With cd1
    .FileName = "ProxyList"
    If IsFar = 1 Then .Filter = "ÝÇíá Ê˜ÓÊ | *.txt"
    If IsFar = 0 Then .Filter = "Text File | *.txt"
    If IsFar = 1 Then .DialogTitle = "áØÝÇ ãÍá ÐÎíÑå ÝÇíá ÑÇ ÊÚííä ˜äíÏ"
    If IsFar = 0 Then .DialogTitle = "Export "
    .ShowSave
End With

If cd1.CancelError = True Then Exit Sub



Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 115)


Open cd1.FileName For Output As #1
    With Adodc1
        .ConnectionString = dbPath
        .RecordSource = "select * from proxy"
        .Refresh
        Bar1.Max = .Recordset.RecordCount
        Do While .Recordset.EOF = False
            st = ""
            st = .Recordset.Fields(1) & ":"
            st = st & .Recordset.Fields(2) & vbTab
            st = st & .Recordset.Fields(3)
            Print #1, st
            .Recordset.MoveNext
            Bar1.Value = Bar1.Value + 1
        Loop
        .Recordset.ActiveConnection = Nothing
    End With
Close #1

Bar1.Value = 0
Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 116)



End Sub

Private Sub mnuFarsi_Click()

    If mnuFarsi.Checked = False Then
        IsFar = 1
        mnuFarsi.Checked = True
        mnuEng.Checked = False
        SaveSetting "ProxyChecker", "Settings", "LANG", IsFar
        ChangeLang
    End If
    
    

End Sub

Private Sub mnuFile_Click()
    Status1.Panels(IsFar * 1 + 1).Text = Split(mnuFile.Caption, "&")(1)
End Sub

Private Sub mnuHelp_Click()
    Status1.Panels(IsFar * 1 + 1).Text = Split(mnuHelp.Caption, "&")(1)
End Sub

Private Sub mnuImport_Click()
    If IsChanged = True Then
        Dim st As String
        st = LoadResString(IsFar * 1000 + 109)
        st = st & vbCrLf & vbCrLf
        st = st & LoadResString(IsFar * 1000 + 113)
        MsgBox st, vbOKOnly, ""
        Exit Sub
    End If
    
With cd1
    .FileName = ""
    If IsFar = 1 Then .Filter = "ÝÇíá Ê˜ÓÊ | *.txt"
    If IsFar = 0 Then .Filter = "Text File | *.txt"
    If IsFar = 1 Then .DialogTitle = "áØÝÇ ãÍá  ÝÇíá ÑÇ ÊÚííä ˜äíÏ"
    If IsFar = 0 Then .DialogTitle = "Import "
    .ShowSave
End With

If cd1.CancelError = True Then Exit Sub
If cd1.FileName = "" Then Exit Sub

frmImport.Run cd1.FileName

End Sub

Private Sub mnuSave_Click()
'
IsChanged = False

With Adodc1
    .ConnectionString = dbPath
    Bar1.Max = lsv1.ListItems.Count
    For i = 1 To lsv1.ListItems.Count
        
            .RecordSource = "select * from proxy where key = " & Val(lsv1.ListItems(i).Text)
            .Refresh
            If lsv1.ListItems(i).SubItems(4) <> "" Then .Recordset.Fields(5) = lsv1.ListItems(i).SubItems(4)
            If lsv1.ListItems(i).SmallIcon = 1 Then
                    .Recordset.Fields(6) = True
                Else
                    .Recordset.Fields(6) = False
            End If
                
            .Recordset.Update
            .Recordset.ActiveConnection = Nothing
        
        Bar1.Value = i
            
            Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 111)

    Next i
End With

Bar1.Value = 0

Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 112)

End Sub

Private Sub mnuSetting_Click()
frmSettings.Show vbModal
End Sub

Private Sub mnuStart_Click()

    IsChanged = True

    mnuStart.Enabled = False
    mnuStop.Enabled = True

    IsScannig = True
     
    Dim TimeOut As Integer
    Dim st As String
    
    st = GetSetting("ProxyChecker", "Settings", "TimeOut")
    
    If st = "" Then TimeOut = 15 Else TimeOut = Val(st)
    
    If lsv1.ListItems.Count <> 0 Then Bar1.Max = lsv1.ListItems.Count
        
    For i = 1 To 100
        If i > lsv1.ListItems.Count Then GoTo Enough
        With Sock1(i)
            .Close
            lsv1.ListItems(i).SmallIcon = 3
            .Tag = i
            .RemoteHost = lsv1.ListItems(i).SubItems(1)
            .RemotePort = lsv1.ListItems(i).SubItems(2)
            .Connect
        End With
        With Timer1(i)
            .Interval = TimeOut * 1000
            .Enabled = True
        End With
        LastRow = i
    Next i
    
Enough:
    
    TimerBar.Enabled = True
    
End Sub

Private Sub mnuStop_Click()

    For i = 1 To 100
        Timer1(i).Enabled = False
        Sock1(i).Close
    Next i

    With lsv1
        For j = 1 To .ListItems.Count
            If .ListItems(j).SmallIcon = 3 Then
                .ListItems(j).SmallIcon = 2
                .ListItems(j).ListSubItems(5).Text = LoadResString(IsFar * 1000 + 102)
            End If
        Next j
    End With
    
    CompleteScan
    
End Sub

Private Sub mnuTools_Click()
    Status1.Panels(IsFar * 1 + 1).Text = Split(mnuTools.Caption, "&")(1)
End Sub

Private Sub Sock1_Connect(Index As Integer)

    Dim Packet As String
    
    Timer1(Index).Enabled = False
    
    Packet = "GET http://login.yahoo.com HTTP/1.0" & vbCrLf
    Packet = Packet & "Accept: */*" & vbCrLf
    Packet = Packet & "Accept-Language: en-us" & vbCrLf
    Packet = Packet & "Connection: Keep-Alive" & vbCrLf & vbCrLf
    
    Sock1(Index).SendData Packet
    
    Timer1(Index).Enabled = True

End Sub

Private Sub Sock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim ReturnData(1 To 100) As String
    
    Timer1(Index).Enabled = False
    
    Sock1(Index).GetData ReturnData(Index), vbString, bytesTotal
    
    If InStr(2, LCase$(ReturnData(Index)), "<title>sign", vbBinaryCompare) <> 0 Then
        
            With lsv1.ListItems(Val(Sock1(Index).Tag))
                .SmallIcon = 1
                .ListSubItems(4).Text = Now
                .ListSubItems(5).Text = LoadResString(IsFar * 1000 + 101)
            End With
        Else
            With lsv1.ListItems(Val(Sock1(Index).Tag))
                .SmallIcon = 2
                .ListSubItems(5).Text = LoadResString(IsFar * 1000 + 102)
            End With
    End If
    
    With Sock1(Index)
        .Close
        If LastRow = lsv1.ListItems.Count Then
                CompleteScan
                Exit Sub
            Else
                Bar1.Value = Bar1.Value + 1
                lsv1.ListItems(LastRow + 1).SmallIcon = 3
                .Tag = LastRow + 1
                .RemoteHost = lsv1.ListItems(LastRow + 1).SubItems(1)
                .RemotePort = lsv1.ListItems(LastRow + 1).SubItems(2)
                .Connect
                LastRow = LastRow + 1
                Timer1(Index).Enabled = True
        End If
    End With

End Sub

Private Sub Sock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    Timer1(Index).Enabled = False
    
    
    With lsv1.ListItems(Val(Sock1(Index).Tag))
        .SmallIcon = 2
        .ListSubItems(5).Text = LoadResString(IsFar * 1000 + 102)
    End With
        
    With Sock1(Index)
        .Close
        If LastRow = lsv1.ListItems.Count Then
                CompleteScan
                Exit Sub
            Else
                Bar1.Value = Bar1.Value + 1
                lsv1.ListItems(LastRow + 1).SmallIcon = 3
                .Tag = LastRow + 1
                .RemoteHost = lsv1.ListItems(LastRow + 1).SubItems(1)
                .RemotePort = lsv1.ListItems(LastRow + 1).SubItems(2)
                .Connect
                LastRow = LastRow + 1
                Timer1(Index).Enabled = True
        End If
    End With
    
End Sub



Private Sub Timer1_Timer(Index As Integer)

    Timer1(Index).Enabled = False
    
    
    With lsv1.ListItems(Val(Sock1(Index).Tag))
        .SmallIcon = 2
        .ListSubItems(5).Text = LoadResString(IsFar * 1000 + 102)
    End With
        
    With Sock1(Index)
        .Close
        If LastRow = lsv1.ListItems.Count Then
                CompleteScan
                Exit Sub
            Else
                Bar1.Value = Bar1.Value + 1
                lsv1.ListItems(LastRow + 1).SmallIcon = 3
                .Tag = LastRow + 1
                .RemoteHost = lsv1.ListItems(LastRow + 1).SubItems(1)
                .RemotePort = lsv1.ListItems(LastRow + 1).SubItems(2)
                .Connect
                LastRow = LastRow + 1
                Timer1(Index).Enabled = True
        End If
    End With
    


End Sub

Private Sub TimerBar_Timer()

            Dim st As String
            For i = 0 To Val(Right(CStr(Bar1.Value), 1))
                st = st & " ."
            Next i
            Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 107) & st


End Sub

Private Sub TimerList_Timer()

TimerList.Enabled = False

    lsv1.ListItems.Clear

    Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 103)

    With Adodc1
        .ConnectionString = dbPath
        .RecordSource = "select * from proxy "
        .Refresh
        If .Recordset.RecordCount <> 0 Then
            Bar1.Max = .Recordset.RecordCount
            Dim i As Integer
            Dim imgindex As Integer
            Dim StatusStr As String
            i = 1
            Do While .Recordset.EOF = False
                If .Recordset.Fields(6) = True Then
                        imgindex = 1
                        StatusStr = LoadResString(IsFar * 1000 + 101)
                    Else
                        imgindex = 2
                        StatusStr = LoadResString(IsFar * 1000 + 102)
                End If
                lsv1.ListItems.Add i, , .Recordset.Fields(0), , imgindex
                lsv1.ListItems(i).ListSubItems.Add 1, , .Recordset.Fields(1)
                lsv1.ListItems(i).ListSubItems.Add 2, , .Recordset.Fields(2)
                lsv1.ListItems(i).ListSubItems.Add 3, , .Recordset.Fields(3)
                If IsNull(.Recordset.Fields(5)) = True Then
                        lsv1.ListItems(i).ListSubItems.Add 4, , ""
                    Else
                        lsv1.ListItems(i).ListSubItems.Add 4, , .Recordset.Fields(5)
                End If
                lsv1.ListItems(i).ListSubItems.Add 5, , StatusStr
                Bar1.Value = Bar1.Value + 1
                i = i + 1
                .Recordset.MoveNext
            Loop
            .Recordset.ActiveConnection = Nothing
            Bar1.Value = 0
        End If
    End With
    
    Status1.Panels(IsFar * 1 + 1).Text = LoadResString(IsFar * 1000 + 104)
    
    If lsv1.ListItems.Count = 0 Then
            Exit Sub
        Else
            If isRunnig = False Then
                For i = 1 To 100
                    Load Sock1(i)
                    Load Timer1(i)
                Next i
            End If
    End If
    
    isRunnig = True
    
    

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1
            mnuStart_Click
        Case 2
            mnuStop_Click
        Case 4
            mnuAdd_Click
        Case 5
            mnuEditProxy_Click
        Case 6
            mnuDel_Click
        Case 8
            mnuSetting_Click
        Case 9
            If IsFar = 0 Then mnuFarsi_Click Else mnuEng_Click
        Case 11
            mnuHelp_Click
    End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then mnuEng_Click Else mnuFarsi_Click
End Sub
