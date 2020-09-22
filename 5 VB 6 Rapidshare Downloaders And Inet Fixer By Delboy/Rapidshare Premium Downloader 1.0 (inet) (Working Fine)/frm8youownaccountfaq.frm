VERSION 5.00
Begin VB.Form frm8youownaccountfaq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "None......"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2595
   Icon            =   "frm8youownaccountfaq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "No FAQ YET..."
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
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

