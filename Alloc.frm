VERSION 5.00
Begin VB.Form frmAlloc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Size"
   ClientHeight    =   2610
   ClientLeft      =   1575
   ClientTop       =   4860
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "Alloc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAlloc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "1"
      Top             =   1455
      Width           =   735
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MBytes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   6
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MBytes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblFreeDisk 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Capture File Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Free Space"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Amount of disk space to set aside for the capture file."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3420
   End
End
Attribute VB_Name = "frmAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private available As Long

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.left = "2500"
cmdCancel.top = "2190"
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.left = "2400"
cmdCancel.top = "2160"
End Sub

Private Sub cmdOK_Click()
Call capFileAlloc(frmMain.capwnd, txtAlloc.Text * ONE_MEGABYTE)
Unload Me
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOK.left = "520"
cmdOK.top = "2200"
End Sub

Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdOK.left = "480"
cmdOK.top = "2160"
End Sub

Private Sub Form_Load()
Dim capfilesize As Long
Dim path As String

On Error Resume Next 'if user has deleted file this is necessary
path = capFileGetCaptureFile(frmMain.capwnd)
path = left$(path, 3)
lblFreeDisk.Caption = vbGetAvailableMBytesAsString(path) 'use GetFree.bas to handle large ( > 2GB ) volumes...

capfilesize = FileLen(capFileGetCaptureFile(frmMain.capwnd))
If capfilesize > (ONE_MEGABYTE / 2) Then
    txtAlloc.Text = capfilesize / ONE_MEGABYTE
Else
    txtAlloc.Text = 1
End If
txtAlloc.SelStart = 0
txtAlloc.SelLength = Len(txtAlloc.Text)

End Sub



Private Sub txtAlloc_Change()
If Val(txtAlloc.Text) < 0 Then txtAlloc.Text = 1
If Val(lblFreeDisk.Caption) < 1 Then Exit Sub
If Val(txtAlloc.Text) > Val(lblFreeDisk.Caption) Then txtAlloc.Text = lblFreeDisk.Caption
End Sub

Private Sub txtAlloc_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub
