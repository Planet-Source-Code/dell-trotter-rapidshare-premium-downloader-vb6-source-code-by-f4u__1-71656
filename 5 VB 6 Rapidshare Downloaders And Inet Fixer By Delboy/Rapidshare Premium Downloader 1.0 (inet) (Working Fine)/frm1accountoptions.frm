VERSION 5.00
Begin VB.Form frm1accountoptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.0 By F4U"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm1accountoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "FAQ (NOT ADDED YET)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use my own account"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Use the built in account"
      Height          =   435
      Left            =   2400
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      Picture         =   "frm1accountoptions.frx":0CCA
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Created By Delboy With Help From lintz"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "www.freesoftwarealliance.com"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   3615
   End
End
Attribute VB_Name = "frm1accountoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
Command5.Enabled = False
End Sub

Private Sub Command1_Click()
frm3mainyourown.Show
Unload Me
End Sub

Private Sub Command2_Click()
frm4about.Show
End Sub

Private Sub Command3_Click()
frm2mainbuiltin.Show
Unload Me
End Sub

Private Sub Command4_Click()
frmupdater.Show
End Sub

Private Sub Command5_Click()
frm6faq.Show
End Sub
