VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2340
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1615.11
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4800
      Picture         =   "frmAbout.frx":058A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Email  :   kargar.reza@gmail.com"
      Height          =   225
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblAbout 
      Alignment       =   1  'Right Justify
      Caption         =   "»—‰«„Â ‰ÊÌ”   :  —÷« ò«—ê—                 ·›‰ : 09122767401"
      Height          =   225
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1107.8
      Y2              =   1107.8
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Proxy Checker"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   570
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1118.153
      Y2              =   1118.153
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me

End Sub

Private Sub Form_Load()

    With Me
        .Caption = LoadResString(IsFar * 1000 + 137)
        .RightToLeft = CBool(IsFar)
    End With

    
    With lblAbout
        .Caption = LoadResString(IsFar * 1000 + 169)
        .RightToLeft = CBool(IsFar)
    End With
    
    cmdOK.Caption = LoadResString(IsFar * 1000 + 160)

End Sub
