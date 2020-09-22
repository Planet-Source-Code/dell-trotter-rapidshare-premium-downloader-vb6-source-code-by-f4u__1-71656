VERSION 5.00
Begin VB.Form frm7builtinaccountfaq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "None..."
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2475
   Icon            =   "frm7builtinaccountfaq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2475
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "No FAQ YET..."
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm7builtinaccountfaq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

