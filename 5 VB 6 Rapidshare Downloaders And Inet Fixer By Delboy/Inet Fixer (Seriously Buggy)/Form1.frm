VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIX INET BUG"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here To Fix Inet Error"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Error
'Copys The Stuff, box 1 is were it goes from box 2 is destination
FileCopy App.Path & "\" & "MSINET.OCX", Drive1.Drive & "\Windows\System32\MSINET.OCX"

MsgBox "Inet Fixing Successful"
Exit Sub
Error:
MsgBox "Random Error Has Accured, Please Restart"
End Sub
