VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm2mainbuiltinsingle 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.2 Beta - Built In Account Single Mode By Delboy {F4U}"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   Icon            =   "frm2mainbuiltinsingle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Single Download Mode"
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   9135
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download"
         Height          =   375
         Left            =   7440
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdrotate 
         Caption         =   "Rotate Account"
         Height          =   255
         Left            =   7440
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtdownloader 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Text            =   "Download Link"
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "File Name E.g Hello.rar"
         Top             =   840
         Width           =   2415
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   8895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   3480
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Download Link:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Caption         =   "File Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Created By Delboy With Help From lintz"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   8895
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "www.freesoftwarealliance.com"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   8895
      End
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "cmdlogin"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   8760
      Width           =   2415
   End
   Begin VB.TextBox txtuser 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Text            =   "txtuser"
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton cmdaccounts 
      Caption         =   "cmdaccounts"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox txtpass 
      Height          =   405
      Left            =   3600
      TabIndex        =   3
      Text            =   "txtpass"
      Top             =   8160
      Width           =   2415
   End
   Begin VB.ListBox Pass 
      Height          =   1230
      Left            =   1800
      TabIndex        =   2
      Top             =   7800
      Width           =   1695
   End
   Begin VB.ListBox User 
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   120
      Picture         =   "frm2mainbuiltinsingle.frx":0CCA
      ScaleHeight     =   3195
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6360
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frm2mainbuiltinsingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
ProgressBar1.Visible = False
cmdDownload.Enabled = False
Call cmdaccounts_Click
Call cmdrotate_Click
Call cmdlogin_Click
End Sub
Private Sub cmdrotate_Click()
If User.ListIndex = User.ListCount - 1 Then
    User.ListIndex = 0 - 1
End If
User.ListIndex = User.ListIndex + 1
txtuser.Text = User.List(User.ListIndex)

If Pass.ListIndex = Pass.ListCount - 1 Then
    Pass.ListIndex = 0 - 1
End If
Pass.ListIndex = Pass.ListIndex + 1
txtpass.Text = Pass.List(Pass.ListIndex)
Call cmdlogin_Click
End Sub

Private Sub cmdaccounts_Click()
' http://rapidshare.com/cgi-bin/premium.cgi?accountid=AccountID&password=PassWord&premiumlogin=1
User.AddItem "ACCOUNT1"
Pass.AddItem "PASS1"
End Sub


Private Sub cmdDownload_Click()
Screen.MousePointer = vbHourglass

ProgressBar1.Value = 0

ProgressBar1.Visible = True 'show progressbar

'This downloads the file and saves to your machine
DownloadFile txtdownloader.Text, Dir1.path & "\" & txtFileName.Text

Screen.MousePointer = vbDefault
MsgBox "Download Complete"

ProgressBar1.Visible = False

End Sub

Private Sub cmdlogin_Click()
Inet1.OpenURL ("http://rapidshare.com/cgi-bin/premium.cgi?accountid=" & txtuser.Text & "&password=" & txtpass.Text & "&premiumlogin=1")
cmdDownload.Enabled = True
End Sub

Sub DownloadProgress(intPercent As String)
    ProgressBar1.Value = intPercent ' Update file download progress
End Sub


'Public Function DownloadFile(strURL As String, strDestination As String) As Boolean
Public Sub DownloadFile(strURL As String, strDestination As String) 'As Boolean
Const CHUNK_SIZE As Long = 1024
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim strHeader As String
Dim b() As Byte
Dim i As Integer

DoEvents
    
With Inet1
    
.URL = strURL
.Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
        
While .StillExecuting
DoEvents
Wend

strHeader = .GetHeader
End With
    
    
strHeader = Inet1.GetHeader("Content-Length")
lngFileLength = Val(strHeader)

DoEvents
    
lngBytesReceived = 0

intFile = FreeFile()

Open strDestination For Binary Access Write As #intFile

Do
b = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
Put #intFile, , b
lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1

DownloadProgress (Round((lngBytesReceived / lngFileLength) * 100))
DoEvents
Loop While UBound(b, 1) > 0

Close #intFile
 
End Sub


Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub
