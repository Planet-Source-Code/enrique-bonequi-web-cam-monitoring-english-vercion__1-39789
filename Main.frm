VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SCM (Security Camera Monitoring)"
   ClientHeight    =   4905
   ClientLeft      =   2850
   ClientTop       =   3405
   ClientWidth     =   6315
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   Begin MSComDlg.CommonDialog mmon 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16776960
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1E16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   900
      ButtonWidth     =   767
      ButtonHeight    =   741
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "abrir"
            Object.ToolTipText     =   "Abrir Video Del Disco Duro..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "destino"
            Object.ToolTipText     =   "Destino Del Video"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "video"
            Object.ToolTipText     =   "Capturar Secuencia De Video"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "frame"
            Object.ToolTipText     =   "Capturar Frame Sola"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ventana"
            Object.ToolTipText     =   "Abrir Otra Ventana De Monitoreo"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   1920
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSetCapFile 
         Caption         =   "&Set Capture File"
      End
      Begin VB.Menu mnuAllocFileSpace 
         Caption         =   "&File Size"
      End
      Begin VB.Menu mnuspacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveFileAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuSaveFrame 
         Caption         =   "Save Frame As..."
      End
      Begin VB.Menu mnuspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucolordefondo 
         Caption         =   "&BackColor..."
      End
      Begin VB.Menu mnulinea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAudioFmt 
         Caption         =   "&Audio Format"
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "&Format..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSource 
         Caption         =   "S&ource..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Video"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "&Compression..."
      End
      Begin VB.Menu mnuspacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOverlay 
         Caption         =   "&Overlay"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDriver 
         Caption         =   "<none>"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "&Capture"
      Begin VB.Menu mnuCapFrame 
         Caption         =   "Single Frame"
      End
      Begin VB.Menu mnuCapFrames 
         Caption         =   "&Frames..."
      End
      Begin VB.Menu mnuCapVid 
         Caption         =   "&Video..."
      End
   End
   Begin VB.Menu mnumoni 
      Caption         =   "&Monitoring"
      Begin VB.Menu mnumonitoring 
         Caption         =   "&Open Other Video Monitoring Window"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSysInfo 
         Caption         =   "System Info"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hCapWnd As Long       ' Handle to the Capture Windows
Private nDriverIndex As Long  ' video driver index (default 0)
Private m_CapParams As CAPTUREPARMS
'Public property to prevent reentrancy in Form_Resize event
Public AutoSizing As Boolean
'read-only public property to allow other forms to retrieve hwnd of Cap Window
Public Property Get capwnd() As Long
    capwnd = hCapWnd
End Property
'read-only properties for sizing
Public Property Get MenuHeight() As Long
    MenuHeight = GetSystemMetrics(SM_CYMENU)
End Property
Public Property Get CaptionHeight() As Long
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
End Property
Public Property Get XBorder() As Long
    If Me.Appearance = 0 Then   'flat
        XBorder = GetSystemMetrics(SM_CXBORDER)
    Else                        '3D
        XBorder = GetSystemMetrics(SM_CXEDGE)
    End If
End Property
Public Property Get YBorder() As Long
    If Me.Appearance = 0 Then   'flat
        YBorder = GetSystemMetrics(SM_CYBORDER)
    Else                        '3D
        YBorder = GetSystemMetrics(SM_CYEDGE)
    End If
End Property


Private Sub Form_Load()
    Dim retVal As Boolean
    Dim numDevs As Long
    Dim left As Long, top As Long
    
    'load trivial settings first
    Me.BackColor = Val(GetSetting(App.Title, "preferences", "backcolor", "&H00FFFFFF&")) 'default to dk gray
    On Error Resume Next
    left = (Screen.Width - Me.Width) / 2 'center window by default
    top = (Screen.Height - Me.Height) / 2
    On Error GoTo 0
    left = Val(GetSetting(App.Title, "preferences", "left", left))
    top = Val(GetSetting(App.Title, "preferences", "top", top))
    If left < 0 Then left = 0 'just make sure app isn't off the screen
    If top < 0 Then top = 0
    If left > Screen.Width - Me.Width Then left = Screen.Width - Me.Width
    If top > Screen.Height - Me.Height Then top = Screen.Height - Me.Height
    Me.left = left
    Me.top = top
    
    numDevs = VBEnumCapDrivers(Me)
    If 0 = numDevs Then
        MsgBox "capture hardware No detected", vbCritical, App.Title
       Exit Sub
    End If
    nDriverIndex = Val(GetSetting(App.Title, "driver", "index", "0"))
    'if invalid entry is in registry use default (0)
    If mnuDriver.UBound < nDriverIndex Then
        nDriverIndex = 0
    End If
    mnuDriver(nDriverIndex).Checked = True
    '//Create Capture Window
    'Call capGetDriverDescription( nDriverIndex,  lpszName, 100, lpszVer, 100  '// Retrieves driver info
    hCapWnd = capCreateCaptureWindow("VB CAP WINDOW", WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Me.hWnd, 0)
    If 0 = hCapWnd Then
        MsgBox "ERROR", vbCritical, App.Title
        Exit Sub
    End If
    retVal = ConnectCapDriver(hCapWnd, nDriverIndex)
    If False = retVal Then
        MsgBox "ERROR", vbInformation, App.Title
    Else
        #If USECALLBACKS = 1 Then
            ' if we have a valid capwnd we can enable our status callback function
            Call capSetCallbackOnStatus(hCapWnd, AddressOf StatusProc)
            Debug.Print "---Callback set on capture status---"
        #End If
    End If
        '// Set the video stream callback function
'    capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
'    capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
 

End Sub


Public Sub Form_Resize()
    
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    'kludgy way to restrict min form size - better way is to subclass MINMAXINFO messages
    If Me.WindowState = vbMinimized Then Exit Sub 'runtime error was happening when user minimized app...
    If Me.ScaleWidth < 320 Then Me.Width = (320 + (Me.XBorder * 2)) * Screen.TwipsPerPixelX
    If Me.ScaleHeight < 240 Then Me.Height = (240 + (Me.YBorder * 2) + Me.MenuHeight + Me.CaptionHeight) * Screen.TwipsPerPixelY
    'Get the capture window attributes
    retVal = capGetStatus(hCapWnd, capStat)
        
    If retVal Then
        'center the capture window on the form
        Call SetWindowPos(hCapWnd, _
                    0&, _
                    (Me.ScaleWidth - capStat.uiImageWidth) / 2, _
                    (Me.ScaleHeight - capStat.uiImageHeight) / 2, _
                    0&, _
                    0&, _
                    SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOSENDCHANGING) 'by telling Windows not to send
                                                                    'WM_WINDOWPOSCHANGING messages we
                                                                    'eliminate the need for a reentrancy flag
    End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)

'save trivial settings
If Me.WindowState = vbDefault Then
    Call SaveSetting(App.Title, "preferences", "left", Me.left)
    Call SaveSetting(App.Title, "preferences", "top", Me.top)
End If

'unsubclass if necessary
#If USECALLBACKS = 1 Then
    ' Disable status callback
    Call capSetCallbackOnStatus(hCapWnd, 0&)
    Debug.Print "---Capture status callback released---"
#End If

'disconnect VFW driver
Call mVFW.capDriverDisconnect(hCapWnd)
'destroy CapWnd
If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd)
End

End Sub


Private Sub mnuAbout_Click()
    Dim AboutWnd As frmAbout
    Set AboutWnd = New frmAbout
    
    AboutWnd.Show vbModal, Me
    
    Set AboutWnd = Nothing
End Sub

Private Sub mnuAllocFileSpace_Click()
    Dim AllocWnd As frmAlloc
    Set AllocWnd = New frmAlloc
    
    AllocWnd.Show vbModal, Me
    
    Set AllocWnd = Nothing

End Sub

Private Sub mnuAudioFmt_Click()
    Call SetAudioFormatDlg(Me.hWnd)
End Sub

Private Sub mnuCapFrame_Click()

    Call capGrabFrame(hCapWnd)

End Sub

Private Sub mnuCapFrames_Click()
    Dim FrameCapWnd As frmCapFrame
    
    Set FrameCapWnd = New frmCapFrame
    FrameCapWnd.Show vbModal, Me
    
    Set FrameCapWnd = Nothing
    
End Sub



Private Sub mnuCapVid_Click()
    Dim retVal As Boolean
    Dim VidCapWnd As frmCapVid
    
    Set VidCapWnd = New frmCapVid
    VidCapWnd.Show vbModal, Me
    If VidCapWnd.Tag <> "" Then 'use tag to indicate whether user has pressed OK or not
'            // Capture video sequence
        retVal = capCaptureSequence(hCapWnd)
        Unload VidCapWnd 'reclaim mem
    End If
    Set VidCapWnd = Nothing
End Sub

Private Sub mnucolordefondo_Click()
C.ShowColor
Me.BackColor = C.Color
End Sub

Private Sub mnuCompression_Click()

    Call capDlgVideoCompression(hCapWnd)

End Sub

Private Sub mnuCopy_Click()
    
    Call capEditCopy(hCapWnd)

End Sub

Private Sub mnuDisplay_Click()

    Call capDlgVideoDisplay(hCapWnd)
    
End Sub

Private Sub mnuDriver_Click(index As Integer)
    Dim retVal As Boolean
    
    retVal = ConnectCapDriver(hCapWnd, index)
    If False = retVal Then
        MsgBox "ERROR", vbInformation, App.Title
    Else
        Call SaveSetting(App.Title, "driver", "index", CStr(index)) 'save selected device index
    End If
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuFormat_Click()

    Call capDlgVideoFormat(hCapWnd)
    Call ResizeCaptureWindow(hCapWnd)

End Sub



Private Sub mnumonitoring_Click()
On Error GoTo errorhandler
       Dim lngresult As Long
       lngresult = Shell(App.path + "/" + App.EXEName)
       Exit Sub
errorhandler:
MsgBox "another video window cannot be opened"
End Sub

Private Sub mnuOverlay_Click()
    
    mnuOverlay.Checked = Not (mnuOverlay.Checked)
    Call capOverlay(hCapWnd, mnuOverlay.Checked)
    
End Sub



Private Sub mnuPreview_Click()

    mnuPreview.Checked = Not (mnuPreview.Checked)
    Call capPreview(hCapWnd, mnuPreview.Checked)

End Sub


Private Sub mnuSaveFileAs_Click()
Dim FileName As String
Dim retVal As Boolean

retVal = VBGetSaveFileNamePreview(FileName, _
                            FileMustExist:=False, _
                            HideReadOnly:=True, _
                            filter:="AVI Files (*.avi)|*.avi", _
                            DefaultExt:="avi", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capFileSaveAs(hCapWnd, FileName)
    If True <> retVal Then
        MsgBox "Problems saving capture file", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSaveFrame_Click()
Dim FileName As String
Dim retVal As Boolean

retVal = VBGetSaveFileName(FileName, _
                            filter:="DIB Bitmap Files (*.bmp)|*.bmp", _
                            DlgTitle:="Save Single Frame", _
                            DefaultExt:="bmp", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capFileSaveDIB(hCapWnd, FileName)
    If True <> retVal Then
        MsgBox "Problems saving frame file", vbInformation, App.Title
    End If
End If
End Sub



Private Sub mnuSetCapFile_Click()
Dim CapFile As String
Dim CapFileTitle As String
Dim CapFileDir As String
Dim retVal As Boolean
Dim nfileLen As Long

CapFile = mVFW.capFileGetCaptureFile(hCapWnd)
CapFileTitle = VBGetFileTitle(CapFile)
CapFileDir = left$(CapFile, Len(CapFile) - Len(CapFileTitle))
retVal = VBGetOpenFileNamePreview(CapFile, _
                            FileTitle:=CapFileTitle, _
                            filter:="AVI Files (*.avi)|*.avi", _
                            InitDir:=CapFileDir, _
                            DlgTitle:="Set Capture File", _
                            FileMustExist:=False, _
                            HideReadOnly:=True, _
                            DefaultExt:="avi", _
                            Owner:=Me.hWnd)
If True = retVal Then 'user did not cancel
    retVal = mVFW.capFileSetCaptureFile(hCapWnd, CapFile)
    If 0 = retVal Then
        MsgBox "ERROR: " & CapFileTitle, vbInformation, App.Title
        Exit Sub
    Else
        'capture file was changed successfully let's allocate some disk space for it
        'but only if it doesn't already exist
        On Error Resume Next
        nfileLen = FileLen(CapFile)
        If Err.Number = 53 Then 'file does not exist
            Call mnuAllocFileSpace_Click
        End If
    End If
End If
End Sub

Private Sub mnuSource_Click()
'   /*
'    * Display the Video Source dialog when "Source" is selected from the
'    * menu bar.
'    */
    
    Call capDlgVideoSource(hCapWnd)

End Sub



Private Sub mnuSysInfo_Click()
Call ShellAbout(Me.hWnd, _
                App.Title & " System Info Window#OS Information:", _
                vbCrLf & _
                "SCM  Copyright(C) 2002 SoftCrisis", _
                Me.Icon)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "destino"
            mnuSetCapFile_Click
            Case "video"
            mnuCapVid_Click
            Case "frame"
            mnuCapFrames_Click
            Case "ventana"
           mnumonitoring_Click
            Case "abrir"
            mmon.ShowOpen
            vpreview.Show
            vpreview.player.FileName = mmon.FileName
                Case "ayuda"
          mnuAbout_Click
            End Select
            

End Sub





