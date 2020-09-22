VERSION 5.00
Begin VB.Form frm4about 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4440
   Icon            =   "frm4about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disclaimer"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frm4about.frx":0CCA
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frm4about.frx":0DCA
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.freesoftwarealliance.com"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   4455
   End
End
Attribute VB_Name = "frm4about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
