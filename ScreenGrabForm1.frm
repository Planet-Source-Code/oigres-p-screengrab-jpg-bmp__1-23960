VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawMode        =   6  'Mask Pen Not
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xStart As Single, yStart As Single, bMouseDown As Boolean
Dim xs, ys

Private Sub Form_Unload(Cancel As Integer)

    MDIForm1.Visible = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'start of mouse down coords xStart:yStart
    xStart = x: yStart = y
    
    bMouseDown = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next

    If bMouseDown = True Then

        movelines x, y
        
        xs = x: ys = y
        Dim xbig, ybig
        'place label info at bottom right;get biggest coords
        xbig = max(xStart, xs)
        ybig = max(yStart, ys)
        
        Label1.Visible = False
        Label1.left = xbig + 4: Label1.top = ybig + 4
        'who saves a 1 pixel line? same start and new x/y pos makes width/height =0
        Dim grabwidth As Long, grabheight As Long
        If xStart = x Then
            grabwidth = 0
        Else
            grabwidth = Abs(xStart - x) + 1
        End If
        
        If yStart = y Then
            grabheight = 0
        Else
            grabheight = Abs(yStart - y) + 1
        End If
        
        Label1.Caption = "X=" & Format$(x, "0000") & vbCrLf & "Y=" & Format$(y, "0000") & vbCrLf & "Width=" & Format$(grabwidth, "0000") _
           & vbCrLf & "Height=" & Format$(grabheight, "0000")
        '
        Label1.Visible = True
    Else
        'show  coords in label
        Label1.Visible = True
        'move label to left if off screen right
        If x + 4 + Label1.width > Screen.width / Screen.TwipsPerPixelX Then
            'MsgBox "labe3l off"
            Label1.left = x - (Label1.width + 4)
        Else
            Label1.left = x + 4
        End If
        'move label up if off bottom of screen
        If y + 4 + Label1.Height > Screen.Height / Screen.TwipsPerPixelY Then
            'MsgBox "labe3l off"
            Label1.top = y - (Label1.Height + 4)
        Else
            Label1.top = y + 4
        End If
        'Label1.top = y + 4
        
        'just shows xy coords
        Label1.Caption = "X=" & Format$(x, "0000") & vbCrLf & "Y=" & Format$(y, "0000")
        
    End If

    'Form1.Caption = "X= " & Format$(X, "0000") & ": Y= " & Format$(Y, "0000")
   ' Form1.Caption = Format$(x, "0000") & ":" & Format$(y, "0000") & ":" & Format$(Abs(x - xStart), "0000") _
    '   & ":" & Format$(Abs(y - yStart), "0000")

End Sub

Private Sub movelines(x As Single, y As Single)

    If Not (xs = 0 And ys = 0) Then

        'delete previous
        '''-Form1.Line (xStart, yStart)-(xs - 1, ys - 1), , B
        Form1.Line (xStart, yStart)-(xs, ys), , B

    End If

    'draw selection square in invert drawmode
    '''-Form1.Line (xStart, yStart)-(x - 1, y - 1), , B
    Form1.Line (xStart, yStart)-(x, y), , B

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo errMouseUp
    ''Shape1.Visible = False
    Label1.Visible = False

    bMouseDown = False
    ''Form1.Line (xStart, yStart)-(xs, ys), , B
    '''Form1.Line (xStart, yStart)-(xs, ys), , B
    'delete previous
    Form1.Line (xStart, yStart)-(xs, ys), , B

    'Form1.Line (xs, 0)-(xs, ys - 10) '(10 + Shape1.Width))
   
    Dim xwidth, yheight
    Dim startx, starty
    Dim endx, endy
    xwidth = Abs(xStart - xs)
    yheight = Abs(yStart - ys)
    If debugme = True Then MsgBox "xStart = " & xStart & "yStart= " & yStart
    If debugme = True Then MsgBox xwidth & ":" & yheight
    'if mouse start x/y positions = new x/y positions
    If xStart = x And yStart = y Then
        If debugme = True Then MsgBox "xStart =x and yStart=y"
        xs = 0: ys = 0
        Unload Me
        'stops rest of code executing
        Exit Sub
    End If
    'get new form to blit to
    If xwidth <= 0 Or yheight <= 0 Then
        MsgBox "No screen grab width or height"
        'possible bug; not initialised variables;9 june 2001
        xs = 0: ys = 0
        Exit Sub
    End If
    'create new child forms of MDI
    Dim frmChild As New frmChild
    frmChild.Show

    If MDIForm1.ActiveForm Is Nothing Then
    'somehow we have no child form
        MsgBox "need form to blit to"
        Exit Sub
    End If

    frmChild.Picture1.Visible = False

    With MDIForm1.ActiveForm.Picture1

        .BackColor = &HFF00FF
        .Cls
        ''
        '.Width = xwidth + 150
        ''.Width = Screen.TwipsPerPixelX * (xwidth + 8)
        .width = xwidth + 1

        If debugme = True Then MsgBox .width

        '''.Width = xwidth 'Shape1.Width
        ''.Height = yheight + 150 'Shape1.Height
        ''.Height = Screen.TwipsPerPixelY * (yheight + 26)
        .Height = yheight + 1

        If debugme = True Then MsgBox .Height

        'systemmetrics 26= caption and menubar;8= 3d borders of form
        '
        MDIForm1.ActiveForm.width = Screen.TwipsPerPixelX * (xwidth + 8 + 2)
        MDIForm1.ActiveForm.Height = Screen.TwipsPerPixelY * (yheight + 26 + 2)
        ''    '     '
        ''get the correct coords;swap if need be
        'draw from top left corner down to right
        If xStart <= xs And yStart <= ys Then

            startx = xStart: starty = yStart

        End If

        ''draw from bottom right to top left
        If xStart > xs And yStart > ys Then
            startx = xs: starty = ys
        End If

        ''from bottom left to top right
        If xStart < xs And yStart > ys Then
            startx = xStart
            starty = yStart - yheight
        End If

        ''from bottom right to top left
        If xStart > xs And yStart < ys Then
            startx = xStart - xwidth
            starty = yStart
        End If
        '''If xStart > xs Then
        'copy from grab screen form (form1) to picture1 on activeform(frmchild)

        If xwidth > 0 And yheight > 0 Then
            MDIForm1.ActiveForm.Picture1.PaintPicture Form1.Picture, 0, 0, , , startx, starty, xwidth + 1, yheight + 1
        End If

        .Visible = True

    End With

    xs = 0: ys = 0
    
    'convert picture
    MDIForm1.ActiveForm.Picture1.Picture = MDIForm1.ActiveForm.Picture1.Image
    frmChild.Picture1.Visible = True
    
    Unload Me
    Exit Sub

errMouseUp:
    xs = 0: ys = 0
    MsgBox Err.Description & ": Error number " & Err.Number

End Sub

'Private Sub savDesktop()
'Dim sI As String
'Dim c As New cDIBSection
'Dim Quality, qual1, errmsg, n, fileSaved
'
'n = 1
'Get_Dsktop_Path
'  c.CreateFromPicture MDIForm1.ActiveForm.Picture1.Picture
'fileSaved = "Screen Capture"
'pathFileName = "Screen Capture.jpg"
'dsktopPath = TrueDesktopPath & "\" & pathFileName
'StartOver: 'If Quick = 1 Then
'If Dir(dsktopPath) = pathFileName Then
'n = n + 1
'pathFileName = fileSaved & n & ".jpg"
''Debug.Print pathFileName
'dsktopPath = TrueDesktopPath & "\" & pathFileName
''Debug.Print dsktopPath
'GoTo StartOver
'End If
'   qual1 = 90
'   sI = TrueDesktopPath & "\" & pathFileName
'   'Debug.Print sI & "     Last"
''BIG BUG!!!-if save to desktop is selected before save as
''then pathfilename is not found
'    If sI <> "" Then
'      If SaveJPG(c, sI, qual1) Then
'         ' OK!
'      Else
'         MsgBox "Failed to save the picture to the file: '" & sI & "'", vbExclamation
'    End If
'   End If
'
'''End
'End Sub
