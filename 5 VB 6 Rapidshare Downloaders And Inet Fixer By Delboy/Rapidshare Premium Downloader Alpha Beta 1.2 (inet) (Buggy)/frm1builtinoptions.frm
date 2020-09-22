VERSION 5.00
Begin VB.Form frm1builtinoptions 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPTIONS..."
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   Icon            =   "frm1builtinoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Use my own account Multi Download And Single Download"
      Height          =   555
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Built In Single Download"
      Height          =   555
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Built In Multi Download And Single Download"
      Height          =   555
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Use my own account Single Download"
      Height          =   555
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frm1builtinoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm3mainyourownsingle.Show
Unload Me
End Sub

Private Sub Command2_Click()
frm3mainyourown.Show
Unload Me
End Sub

Private Sub Command3_Click()
frm2mainbuiltin.Show
Unload Me
End Sub

Private Sub Command4_Click()
frm2mainbuiltinsingle.Show
Unload Me
End Sub

Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
