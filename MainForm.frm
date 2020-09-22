VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   Caption         =   "Internet Connection Check"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   10
   End
   Begin VB.TextBox Text6 
      Height          =   885
      Left            =   120
      TabIndex        =   7
      Text            =   "Connection type is?"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Text            =   "?"
      Top             =   2160
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "?"
      Top             =   1680
      Width           =   1000
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "?"
      Top             =   1200
      Width           =   1000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "?"
      Top             =   720
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Text            =   "?"
      Top             =   240
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   300
      Left            =   3000
      TabIndex        =   1
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Internet Connection"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   2565
   End
   Begin VB.Label ProgressLabel 
      Caption         =   "Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Check For RAS Installed"
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   2850
   End
   Begin VB.Label Label3 
      Caption         =   "Check if Connected to the Internet"
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   2850
   End
   Begin VB.Label Label2 
      Caption         =   "Check For connection by Proxy"
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   2850
   End
   Begin VB.Label Label1 
      Caption         =   "Check For Modem Connection"
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2850
   End
   Begin VB.Label LanConnection 
      Caption         =   "Check For Lan Connection"
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2850
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'START FORM CODE
Option Explicit


Private Sub Command1_Click()
    ProgressLabel.Caption = "Checking for Lan Connection..."
    ProgressBar1 = 1
    Text1 = IsLanConnection()
    ProgressLabel.Caption = "Checking for Modem Connection..."
    ProgressBar1 = ProgressBar1 + 1
    Text2 = IsModemConnection()
    ProgressLabel.Caption = "Checking for Connection Via Proxy..."
    ProgressBar1 = ProgressBar1 + 1
    Text3 = IsProxyConnection()
    ProgressLabel.Caption = "Checking for Any Internet Connection..."
    ProgressBar1 = ProgressBar1 + 1
    Text4 = IsConnected()
    ProgressLabel.Caption = "Checking if RAS is installed..."
    ProgressBar1 = ProgressBar1 + 1
    Text5 = IsRasInstalled()
    ProgressLabel.Caption = "Getting connection type..."
    ProgressBar1 = ProgressBar1 + 1
    Text6 = ConnectionTypeMsg()
    ProgressBar1 = 10

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Dim CopyRightString As String
CopyRightString = "This program is copyright Mr. Ayman Elbanhawy 1/1/2001"
MsgBox CopyRightString, vbOKOnly, "Copyright Message"
End Sub
