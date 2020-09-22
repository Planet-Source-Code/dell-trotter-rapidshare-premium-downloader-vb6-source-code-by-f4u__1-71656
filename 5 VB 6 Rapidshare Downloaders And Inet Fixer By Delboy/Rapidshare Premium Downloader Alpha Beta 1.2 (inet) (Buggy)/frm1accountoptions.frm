VERSION 5.00
Begin VB.Form frm1accountoptions 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.2 Beta"
   ClientHeight    =   4935
   ClientLeft      =   7125
   ClientTop       =   4995
   ClientWidth     =   4695
   Icon            =   "frm1accountoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4695
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FFFF&
      Caption         =   "Your Own Account"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Built In Account"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   4455
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "CMD2"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "CMD1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtblank 
      Height          =   1095
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frm1accountoptions.frx":0CCA
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "FAQ"
      Height          =   195
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   120
      Picture         =   "frm1accountoptions.frx":0CEE
      ScaleHeight     =   3075
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "About"
      Height          =   195
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblfilecheack 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created By Delboy With Help From lintz"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.freesoftwarealliance.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   4695
   End
End
Attribute VB_Name = "frm1accountoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
MsgBox "This BETA WAS SENT TO slerp FOR BETA TESTING ONLY AND IS NOT READY FOR PUBLIC RELEASE."

Dim retval As String
Dim retval2 As String
retval = Dir$(App.path & "/" & "Downloads")      '-- Sets retval as File To cheack for
retval2 = Dir$(App.path & "/" & "linklist.txt")     '-- Sets retval as File To cheack for

If retval = "Downloads" Then                    '-- Checks For File Existance
lblfilecheack.Caption = "Downloads Folder Found"
Else                                        '-- If File Does Not Exist it continues on to generate the key
Call CMD1_Click
End If

If retval2 = "linklist.txt" Then                    '-- Checks For File Existance
lblfilecheack.Caption = "linklist.txt Found"
Else                                        '-- If File Does Not Exist it continues on to generate the key
Call CMD2_Click
End If
End Sub
Private Sub Command6_Click()
frm1builtinoptions.Show
End Sub
Private Sub Command7_Click()
frm1builtinoptions.Show
End Sub
Private Sub Command2_Click()
frm4about.Show
End Sub
Private Sub Command5_Click()
frm6faq.Show
End Sub
Private Sub CMD1_Click()
Dim folder As String, FileName As String, path As String
path = App.path & "/"
folder = "Downloads"
FileName = Dir(path & folder, vbDirectory)
If FileName = "" Then
MkDir path & folder
lblfilecheack.Caption = "Created Downloads Folder"
Exit Sub
End If
Do While FileName <> " "
Debug.Print path & folder
Exit Sub
Loop
End Sub

Private Sub CMD2_Click()
SaveText txtblank, "linklist.txt"
lblfilecheack.Caption = "Created linklist.txt"
End Sub
