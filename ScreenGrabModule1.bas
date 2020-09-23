Attribute VB_Name = "Module1"
Option Explicit

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
        (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
        ipic As IPicture) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As _
        Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long) As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
'Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
        ByVal hObject As Long) As Long
'Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBmp As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, ByVal lpvBits As Long, ByRef lpbi As BITMAPINFO, ByVal uUsage As Long) As Long
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInstance As Long, ByVal Name As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, _
        ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal lScreenDC As Long, ByVal xSrc As Long, _
        ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
        ByVal hDC As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
        lpRect As RECT) As Long
'Public Declare Function MoveToEx& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
'        ByVal y As Long, ByVal lp As Long)
'Public Declare Function LineTo& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal _
 '       y As Long)
'Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Const debugme As Boolean = False

Public Function savepictureRoutine() As Boolean

    Dim cDib As New cDIBSection
    On Error GoTo errDialog
    'start off false
    savepictureRoutine = False

    If MDIForm1.ActiveForm Is Nothing Then Exit Function

    'get filename with commondialog

    MDIForm1.CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn _
       Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    MDIForm1.CommonDialog1.CancelError = True
    MDIForm1.CommonDialog1.Filter = "All Files|*.*|JPEG|*.jpg;*.JPG|Bitmaps|*.bmp"
    MDIForm1.CommonDialog1.DefaultExt = "bmp"
    MDIForm1.CommonDialog1.FileName = "noname"
    '
    MDIForm1.CommonDialog1.ShowSave

    If debugme = True Then MsgBox MDIForm1.CommonDialog1.FileName

    Dim savefilename As String, strext As String

    If MDIForm1.CommonDialog1.FileName <> "" Then '(1)
        'MsgBox MDIForm1.CommonDialog1.FileName
        savefilename = MDIForm1.CommonDialog1.FileName
        'file extension
        strext = UCase$(Mid$(savefilename, Len(savefilename) - 2, 3))
        'select which routine to use bmp or jpeg
        If strext = "JPG" Then '(2)
            'MsgBox "save as jpeg format"
            cDib.CreateFromPicture MDIForm1.ActiveForm.Picture1.Picture
            Dim qsetting As Long

            qsetting = getQsetting()
            If SaveJPG(cDib, MDIForm1.CommonDialog1.FileName, qsetting) Then '(3)
                ' OK!
                savepictureRoutine = True
                'MDIForm1.ActiveForm.Picture1.Picture = MDIForm1.CommonDialog1.FileName
                'MsgBox MDIForm1.CommonDialog1.FileName
                'reload jpg image
                MDIForm1.ActiveForm.Picture1.Picture = LoadPicture(MDIForm1.CommonDialog1.FileName)
                Exit Function
            Else
                MsgBox "Failed to save the picture to the file: '" & MDIForm1.CommonDialog1.FileName & "'", vbExclamation
                'savepictureRoutine = False
                Exit Function
            End If '(3)
        Else 'jpg
               ' MsgBox "save as " & strext
                If strext = "BMP" Then '(4)
                    'MsgBox "save as bitmap format"
                    'bmp format save
                    SavePicture MDIForm1.ActiveForm.Picture1.Picture, MDIForm1.CommonDialog1.FileName
                    savepictureRoutine = True
                    Exit Function
                End If '(4)
            
        End If 'strext = "jpg"(2)
            'savepicure method saves only bitmap (icon if loaded icon file)
            ''        SavePicture MDIForm1.ActiveForm.Picture1.Picture, MDIForm1.CommonDialog1.FileName
            ''savepictureRoutine = True
    Else
    'no filename to save
    End If '<>""(1)

    Exit Function
errDialog:

    End Function

Public Function max(a, b) As Variant
        If a > b Then
        max = a
        Else
        max = b
        End If
        
End Function
'
Private Function getQsetting() As Long
'get quality settings -index 1-10 ->10-100
'menu index 1-10 is multiplied by 10
'find the checked quality setting in menu
Dim vIndex As Long, qsetting As Long
        For vIndex = MDIForm1.mnuQvalue.LBound To MDIForm1.mnuQvalue.UBound
            If MDIForm1.mnuQvalue.Item(vIndex).Checked = True Then
                getQsetting = vIndex * 10 '1->10;7->70
                Exit For
            End If
        Next vIndex
End Function

'minor change optional value now has default value of 0
Public Function GetScreenSnapshot(Optional ByVal hwnd As Long = 0) As IPictureDisp
'GetScreenSnapshot from  www.vb2themax.com FBalena
    Dim targetDC As Long
    Dim hDC As Long
    Dim tempPict As Long
    Dim oldPict As Long
    Dim wndWidth As Long
    Dim wndHeight As Long
    Dim Pic As PICTDESC
    Dim rcWindow As RECT
    Dim guid(3) As Long

    ' provide the right handle for the desktop window

    If hwnd = 0 Then hwnd = GetDesktopWindow

    ' get window's size
    GetWindowRect hwnd, rcWindow
    wndWidth = rcWindow.right - rcWindow.left
    wndHeight = rcWindow.bottom - rcWindow.top
    ' get window's device context
    targetDC = GetWindowDC(hwnd)

    ' create a compatible DC
    hDC = CreateCompatibleDC(targetDC)

    ' create a memory bitmap in the DC just created
    ' the has the size of the window we're capturing
    tempPict = CreateCompatibleBitmap(targetDC, wndWidth, wndHeight)
    oldPict = SelectObject(hDC, tempPict)

    ' copy the screen image into the DC
    BitBlt hDC, 0, 0, wndWidth, wndHeight, targetDC, 0, 0, vbSrcCopy

    ' set the old DC image and release the DC
    tempPict = SelectObject(hDC, oldPict)
    DeleteDC hDC
    ReleaseDC GetDesktopWindow, targetDC

    ' fill the ScreenPic structure

    With Pic

        .cbSize = Len(Pic)
        .pictType = 1           ' means picture
        .hIcon = tempPict
        .hPal = 0           ' (you can omit this of course)

    End With

    ' convert the image to a IpictureDisp object
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect Pic, guid(0), True, GetScreenSnapshot

End Function
