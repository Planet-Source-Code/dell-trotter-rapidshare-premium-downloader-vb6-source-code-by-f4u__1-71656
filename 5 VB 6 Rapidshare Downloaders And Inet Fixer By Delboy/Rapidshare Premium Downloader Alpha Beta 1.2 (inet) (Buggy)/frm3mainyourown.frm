VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3mainyourownsingle 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.2 - Single Download Own Account Mode"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "frm3mainyourown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Single Download Mode"
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   7575
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "File Name E.g Hello.rar"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtdownloader 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   "Download Link"
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox txtpass 
         Height          =   405
         Left            =   3840
         TabIndex        =   8
         Text            =   "Password"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtuser 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "Username"
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdlogin 
         Caption         =   "Login"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download"
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdlogout 
         Caption         =   "Logout"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   720
         Width           =   3615
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   1560
         Width           =   2775
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   7335
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6840
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "File Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Created By Delboy With Help From lintz"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   7335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "www.freesoftwarealliance.com"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   7335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   240
      Picture         =   "frm3mainyourown.frx":0CCA
      ScaleHeight     =   3075
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frm3mainyourownsingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
ProgressBar1.Visible = False
cmdDownload.Enabled = False
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
Private Sub cmdlogout_Click()
Inet1.OpenURL ("http://rapidshare.com/cgi-bin/premium.cgi?logout=1")
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
