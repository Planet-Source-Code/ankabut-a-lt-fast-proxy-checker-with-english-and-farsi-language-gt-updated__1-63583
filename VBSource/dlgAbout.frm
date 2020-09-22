VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IP Text Box"
   ClientHeight    =   1830
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5160
   ClipControls    =   0   'False
   Icon            =   "dlgAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1263.099
   ScaleMode       =   0  'User
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "dlgAbout.frx":058A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   360
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Caption         =   "Email  :    kargar.reza@gmail.com"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   2
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "IP Text Box          By Reza kargar"
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   360
      Width           =   3885
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub
