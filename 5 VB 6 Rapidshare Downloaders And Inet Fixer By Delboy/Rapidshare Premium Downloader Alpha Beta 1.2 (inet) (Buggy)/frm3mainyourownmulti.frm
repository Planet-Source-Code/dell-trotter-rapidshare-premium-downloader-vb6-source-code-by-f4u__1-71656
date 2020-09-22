VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3mainyourown 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rapidshare Downloader 1.2 - Multi Download Your Own Account Mode By Delboy {F4U}"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "frm3mainyourownmulti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   1680
      Picture         =   "frm3mainyourownmulti.frx":0CCA
      ScaleHeight     =   3075
      ScaleWidth      =   7275
      TabIndex        =   25
      Top             =   120
      Width           =   7335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000C000&
      Caption         =   "Account Info"
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   9975
      Begin VB.TextBox txtpass 
         Height          =   405
         Left            =   5280
         TabIndex        =   24
         Text            =   "Password"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtuser 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Text            =   "Username"
         Top             =   240
         Width           =   5055
      End
      Begin VB.CommandButton cmdlogin 
         Caption         =   "Login"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton cmdlogout 
         Caption         =   "Logout"
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Multi Downloader"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   9975
      Begin VB.TextBox txtextralink 
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Text            =   "Extra Link Here"
         Top             =   3720
         Width           =   7455
      End
      Begin VB.CommandButton cmdloadlist 
         Caption         =   "Load The Link List"
         Height          =   255
         Left            =   6840
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmdaddmore 
         Caption         =   "Add Another Link To The List"
         Height          =   255
         Left            =   6840
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton cmdsavelist 
         Caption         =   "Save The Link List"
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton cmdNextItem 
         Caption         =   "Move Onto Next Item (If It Does Not Do It Auto)"
         Height          =   495
         Left            =   6840
         TabIndex        =   7
         Top             =   2040
         Width           =   3015
      End
      Begin VB.ListBox lstlinklist 
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6615
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download"
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton cmdrotate 
         Caption         =   "Rotate Account"
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtdownloader 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "URL To Current File That is Downloading"
         Top             =   4080
         Width           =   8055
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   7680
         TabIndex        =   2
         Text            =   "Hello.rar"
         Top             =   3000
         Width           =   2175
      End
      Begin VB.OptionButton optmanname 
         BackColor       =   &H0000C000&
         Caption         =   "Manual Name"
         Height          =   255
         Left            =   8280
         TabIndex        =   1
         Top             =   2640
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4440
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblstatus 
         BackColor       =   &H0000C000&
         Caption         =   "Status: Noting To Report Yet......................................"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Width           =   9735
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C000&
         Caption         =   "File Name:"
         Height          =   255
         Left            =   6840
         TabIndex        =   18
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Created By Delboy With Help From lintz"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5160
         Width           =   9735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "www.freesoftwarealliance.com"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   5400
         Width           =   9735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C000&
         Caption         =   "Currently Downloading:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000C000&
         Caption         =   "Link To Add To Download List:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label lblnamestatus 
         BackColor       =   &H0000C000&
         Caption         =   "You Have Not Selected A Naming Option."
         Height          =   255
         Left            =   6840
         TabIndex        =   13
         Top             =   3360
         Width           =   3015
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frm3mainyourown"
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
Private Sub cmdlogout_Click()
Inet1.OpenURL ("http://rapidshare.com/cgi-bin/premium.cgi?logout=1")
End Sub


Private Sub cmdaddmore_Click()
lstlinklist.AddItem txtextralink.Text
End Sub

Private Sub cmdloadlist_Click()
Call Loadlistbox(App.path & "/linklist.txt", lstlinklist)
End Sub

Private Sub cmdNextItem_Click()
If lstlinklist.ListIndex = lstlinklist.ListCount - 1 Then
lstlinklist.ListIndex = 0 - 1
MsgBox "Downloading Has Finished Going Back To Item 1"
lblstatus.Caption = "Downloading Has Finished Going Back To Item 1"
    End If
lstlinklist.ListIndex = lstlinklist.ListIndex + 1
txtdownloader.Text = lstlinklist.List(lstlinklist.ListIndex)
' txtFileName.Text = lstlinklist.List(lstlinklist.ListIndex) ' Auto Name By Default
If cmdDownload.Enabled Then
Call cmdDownload_Click
Else
Do
Loop
End If
End Sub

Private Sub cmdsavelist_Click()
Call SaveListBox(App.path & "/linklist.txt", lstlinklist)
End Sub

Private Sub cmdDownload_Click()
Screen.MousePointer = vbHourglass

ProgressBar1.Value = 0

ProgressBar1.Visible = True 'show progressbar

'This downloads the file and saves to your machine
DownloadFile txtdownloader.Text, App.path & "\Downloads" & "\" & txtFileName.Text

Screen.MousePointer = vbDefault
'MsgBox "Download Complete"
Call cmdNextItem_Click
lblstatus.Caption = "Download Complete moving onto next link"
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


Private Sub optautoname_Click()
txtFileName.Text = lstlinklist.List(lstlinklist.ListIndex)
lblnamestatus.Caption = "Autoname Is ON"
End Sub

Private Sub optmanname_Click()
lblnamestatus.Caption = "Manual Name Is ON, Type The RAR Name In The Box"
End Sub

