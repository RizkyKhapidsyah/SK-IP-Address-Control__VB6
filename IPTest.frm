VERSION 5.00
Object = "*\AAriadIPAddressControl.vbp"
Begin VB.Form frmIPAddressTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IP Address Sample"
   ClientHeight    =   1815
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   Begin IPAddress.AriadIPAddress ipaSample 
      Height          =   330
      Left            =   1170
      TabIndex        =   4
      Top             =   135
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IPAddress       =   ""
   End
   Begin VB.CheckBox chkDisabled 
      Caption         =   "Use Disabled &BackColor"
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   855
      Value           =   1  'Checked
      Width           =   2400
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "&Enabled"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IP &Address:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   810
   End
End
Attribute VB_Name = "frmIPAddressTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Private Sub chkDisabled_Click()
 ipaSample.UseDisabledBackColor = chkDisabled
End Sub

Private Sub chkEnabled_Click()
 ipaSample.Enabled = chkEnabled
End Sub


Private Sub cmdExit_Click()
 Unload Me
End Sub




