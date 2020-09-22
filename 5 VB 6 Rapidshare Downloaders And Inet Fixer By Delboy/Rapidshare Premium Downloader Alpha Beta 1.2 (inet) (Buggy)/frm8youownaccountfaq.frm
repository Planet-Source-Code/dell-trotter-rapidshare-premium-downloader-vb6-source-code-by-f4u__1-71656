VERSION 5.00
Begin VB.Form frm8youownaccountfaq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.2 - Your Own Account FAQ By F4U"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "frm8youownaccountfaq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm8youownaccountfaq.frx":0CCA
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frm8youownaccountfaq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

