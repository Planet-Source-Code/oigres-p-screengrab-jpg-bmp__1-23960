VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Screen Grab"
   ClientHeight    =   3405
   ClientLeft      =   165
   ClientTop       =   5115
   ClientWidth     =   6195
   LinkTopic       =   "MDIForm1"
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   2760
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCloseChild 
         Caption         =   "&CloseChild"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Picture Properties"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "&Capture"
      Begin VB.Menu mnuGrabScreen 
         Caption         =   "&Grab Screen"
      End
      Begin VB.Menu mnuQuality 
         Caption         =   "&Quality"
         Begin VB.Menu mnuQvalue 
            Caption         =   "10"
            Index           =   1
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "20"
            Index           =   2
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "30"
            Index           =   3
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "40"
            Index           =   4
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "50"
            Index           =   5
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "60"
            Index           =   6
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "70"
            Index           =   7
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "80"
            Index           =   8
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "90"
            Index           =   9
         End
         Begin VB.Menu mnuQvalue 
            Caption         =   "100"
            Index           =   10
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'09 June 2001
'ScreenGrab ver 4
'ScreenGrab by oigres P (oigres@postmaster.co.uk)
'additional help from Neil Crosby
'some functions from www.vbacelerator.com - Steve McMahon
'http://www.vbaccelerator.com/codelib/gfx/savejpeg.zip
'GetScreenSnapshot from  www.vb2themax.com
'limitations: savepicture method only saves as bitmap->.bmp
'but also saves as jpg - quality settings set in menu
'can't save dvd;quick time pics or other strangers
'going to other apps when in capture mode is not recommended
'click down then up on same pixel to grab nothing
'tested on win 95/millenium
'to grab a screen; goto menu 'capture/Grab Screen';wait 5 secs
'then draw out square area to grab;set quality settings if saving as jpg
'Jpeg additions from ijl15.dll available at
'http://developer.intel.com/software/products/perflib/ijl/index.htm
'add ijl15.dll to windows\system\ folder.
'Can now save grab as bmp or jpg.
'(still think snagit from www.techsmith.com is very interesting)
'This prog based on screenhunter (as inspiration)
'
''minor change optional value now has default value of 0
'Function GetScreenSnapshot(Optional ByVal hwnd As Long = 0) As IPictureDisp
''GetScreenSnapshot from  www.vb2themax.com FBalena
'    Dim targetDC As Long
'    Dim hDC As Long
'    Dim tempPict As Long
'    Dim oldPict As Long
'    Dim wndWidth As Long
'    Dim wndHeight As Long
'    Dim Pic As PICTDESC
'    Dim rcWindow As RECT
'    Dim guid(3) As Long
'
'    ' provide the right handle for the desktop window
'
'    If hwnd = 0 Then hwnd = GetDesktopWindow
'
'    ' get window's size
'    GetWindowRect hwnd, rcWindow
'    wndWidth = rcWindow.right - rcWindow.left
'    wndHeight = rcWindow.bottom - rcWindow.top
'    ' get window's device context
'    targetDC = GetWindowDC(hwnd)
'
'    ' create a compatible DC
'    hDC = CreateCompatibleDC(targetDC)
'
'    ' create a memory bitmap in the DC just created
'    ' the has the size of the window we're capturing
'    tempPict = CreateCompatibleBitmap(targetDC, wndWidth, wndHeight)
'    oldPict = SelectObject(hDC, tempPict)
'
'    ' copy the screen image into the DC
'    BitBlt hDC, 0, 0, wndWidth, wndHeight, targetDC, 0, 0, vbSrcCopy
'
'    ' set the old DC image and release the DC
'    tempPict = SelectObject(hDC, oldPict)
'    DeleteDC hDC
'    ReleaseDC GetDesktopWindow, targetDC
'
'    ' fill the ScreenPic structure
'
'    With Pic
'
'        .cbSize = Len(Pic)
'        .pictType = 1           ' means picture
'        .hIcon = tempPict
'        .hPal = 0           ' (you can omit this of course)
'
'    End With
'
'    ' convert the image to a IpictureDisp object
'    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
'    ' we use an array of Long to initialize it faster
'    guid(0) = &H7BF80980
'    guid(1) = &H101ABF32
'    guid(2) = &HAA00BB8B
'    guid(3) = &HAB0C3000
'    ' create the picture,
'    ' return an object reference right into the function result
'    OleCreatePictureIndirect Pic, guid(0), True, GetScreenSnapshot
'
'End Function

Private Sub MDIForm_Load()
    '
    Show
    'default 70 jpg quality;set menu check
    mnuQvalue_Click (7)

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If debugme = True Then MsgBox "unload event"

    If debugme = True Then MsgBox MDIForm1.ActiveForm Is Nothing
    
    Unload Form1
    Unload Me

End Sub


Private Sub mnuAbout_Click()

    MsgBox "Screen Grab by oigres P" & vbCrLf & "Email: oigres@postmaster.co.uk", , "Screen Grab"

End Sub

Private Sub mnuCloseChild_Click()

    If Not (ActiveForm Is Nothing) Then Unload ActiveForm

End Sub

Private Sub mnuExit_Click()
    'child forms take care of themselves
    Unload Form1
    Unload Me
End Sub

Private Sub mnuGrabScreen_Click()
    '
    MDIForm1.Visible = False
    Form1.Visible = False
    'start main grab sequence
    Timer1.Enabled = True

End Sub

Private Sub mnuProperties_Click()

    If MDIForm1.ActiveForm Is Nothing Then Exit Sub

    With MDIForm1.ActiveForm.Picture1
        MsgBox "Picture Width= " & CInt(.ScaleX(.Picture.width, vbHimetric, vbPixels)) & " Pixels" & vbCrLf _
           & "Picture Height= " & CInt(.ScaleY(.Picture.Height, vbHimetric, vbPixels)) & " Pixels", , "Picture Properties"
    End With

End Sub

Private Sub mnuQvalue_Click(Index As Integer)
    Static lastindex As Integer
    'remeber last menu quality setting
    If lastindex = Empty Then
        mnuQvalue.Item(Index).Checked = True
    Else
        mnuQvalue.Item(lastindex).Checked = False
        mnuQvalue.Item(Index).Checked = True
    
    End If
    lastindex = Index

End Sub

Private Sub mnuSave_Click()
    '
    'MsgBox "MDIForm1.ActiveForm.IsDirty=" & MDIForm1.ActiveForm.IsDirty
    If MDIForm1.ActiveForm Is Nothing Then Exit Sub
    'no need to save
    If MDIForm1.ActiveForm.IsDirty = False Then Exit Sub

    If savepictureRoutine = True Then
        'reset IsDirty flag
        MDIForm1.ActiveForm.IsDirty = False
        'update menu
        MDIForm1.mnuSave.Enabled = False
    Else
        'no change of isdirty flag/property
    End If

End Sub

Private Sub Timer1_Timer()

    Static count
    ''    MsgBox count
    'wait a bit then get screen;
    'count 0-4 ;timer on 500 interval
    If count > 3 Then

        Form1.Picture = GetScreenSnapshot(0)
        count = 0
        Timer1.Enabled = False
        Form1.Visible = True

    End If

    count = count + 1

End Sub

