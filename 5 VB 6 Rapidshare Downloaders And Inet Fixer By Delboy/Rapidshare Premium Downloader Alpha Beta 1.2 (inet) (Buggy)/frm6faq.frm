VERSION 5.00
Begin VB.Form frm6faq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.0 - FAQ By Delboy {F4U}"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frm6faq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frm6faq.frx":0CCA
   ScaleHeight     =   4935
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdyoafaq 
      Caption         =   "Your Own Account FAQ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5775
   End
   Begin VB.CommandButton cmdbiafaq 
      Caption         =   "Build In Account FAQ"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      Picture         =   "frm6faq.frx":0D3C
      ScaleHeight     =   3075
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "www.freesoftwarealliance.com"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Created By Delboy With Help From lintz"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Top             =   4200
      Width           =   6015
   End
End
Attribute VB_Name = "frm6faq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub cmdbiafaq_Click()
frm7builtinaccountfaq.Show
End Sub

Private Sub cmdyoafaq_Click()
frm8youownaccountfaq.Show
End Sub

