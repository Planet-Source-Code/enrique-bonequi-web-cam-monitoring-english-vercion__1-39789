VERSION 5.00
Begin VB.Form frmCapFrame 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Frame Capture"
   ClientHeight    =   2085
   ClientLeft      =   2460
   ClientTop       =   3900
   ClientWidth     =   4020
   Icon            =   "CapFrame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Cancel"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label cmdCapture 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C&apture"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblFrames 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 Frames"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1065
      TabIndex        =   2
      Top             =   780
      Width           =   1560
   End
   Begin VB.Label lblCapFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   225
      Left            =   15
      TabIndex        =   1
      Top             =   435
      Width           =   3480
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Capture Images To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "frmCapFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.left = "2660"
cmdCancel.top = "1580"
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.left = "2640"
cmdCancel.top = "1560"
End Sub

Private Sub cmdCapture_Click()
    If capCaptureSingleFrame(frmMain.capwnd) Then
        lblFrames.Caption = Val(lblFrames.Caption) + 1 & " Frames"
        cmdCancel.Caption = "Cerrar"
    Else
        MsgBox "ERROR", App.Title, vbInformation
    End If
End Sub

Private Sub cmdCapture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCapture.left = "500"
cmdCapture.top = "1580"
End Sub

Private Sub cmdCapture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCapture.left = "480"
cmdCapture.top = "1560"
End Sub

Private Sub Form_Load()
lblCapFile.Caption = capFileGetCaptureFile(frmMain.capwnd)
If lblCapFile.Caption = "" Then
    lblCapFile.Caption = "<error: no cap file>"
End If
Call capCaptureSingleFrameOpen(frmMain.capwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call capCaptureSingleFrameClose(frmMain.capwnd)
End Sub
