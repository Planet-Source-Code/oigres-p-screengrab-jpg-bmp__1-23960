Attribute VB_Name = "Helpers"
Rem
Rem
Rem
Rem               INTEL CORPORATION PROPRIETARY INFORMATION
Rem  This software is supplied under the terms of a license agreement or
Rem  nondisclosure agreement with Intel Corporation and may not be copied
Rem  or disclosed except in accordance with the terms of that agreement.
Rem      Copyright (c) 1998 Intel Corporation. All Rights Reserved.
Rem
Rem
Rem  File:
Rem    helpers.bas
Rem
Rem  Purpose:
Rem    Helper functions
Rem
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)


Public Function ShowErrorMsg(ByVal context As String, ByVal code As IJLERR)
  Dim message As String
  
  message = "IJL ERROR: [" & code & "]" & " - " & context
  
  Call MsgBox(message, vbExclamation, "Intel(R) JPEG Library")

End Function


'Public Function ConvertFromRGBA(ByVal rgba As Long)
'
'End Function


Public Function LoadJPG(ByRef cDib As cDIBSection, ByVal sFile As String, ByVal jpg_scale As Long) As Boolean
  Dim jerr As IJLERR
  Dim jcprops As JPEG_CORE_PROPERTIES
  Dim aFile As String
  Dim lJPGWidth As Long
  Dim lJPGHeight As Long
  Dim nChannels As Long

  cDib.CleanUp
  
  jerr = ijlInit(jcprops)
  If jerr = IJL_OK Then
      
    ' Write the filename to the jcprops.JPGFile member:
    aFile = StrConv(sFile, vbFromUnicode)
      
    jcprops.JPGFile = StrPtr(aFile)
      
    ' Read the JPEG file parameters:
    jerr = ijlRead(jcprops, IJL_JFILE_READPARAMS)
    If jerr <> IJL_OK Then
      ' Throw error
      Call ShowErrorMsg("FAILED TO READ IMAGE PARAMS", jerr)
    Else
      ' Get the JPGWidth ...
      lJPGWidth = jcprops.JPGWidth
      ' .. & JPGHeight member values:
      lJPGHeight = jcprops.JPGHeight
      
      Select Case jpg_scale
        Case 1
        Case 2
          lJPGWidth = (lJPGWidth + 1) / 2
          lJPGHeight = (lJPGHeight + 1) / 2
            
        Case 4
          lJPGWidth = (lJPGWidth + 3) / 4
          lJPGHeight = (lJPGHeight + 3) / 4
        Case 8
          lJPGWidth = (lJPGWidth + 7) / 8
          lJPGHeight = (lJPGHeight + 7) / 8
      End Select
            
      If jcprops.JPGChannels = 1 Then
        jcprops.JPGColor = IJL_G
        jcprops.DIBColor = IJL_BGR
        nChannels = 3
      ElseIf jcprops.JPGChannels = 3 Then
        jcprops.JPGColor = IJL_YCBCR
        jcprops.DIBColor = IJL_BGR
        nChannels = 3
      ElseIf jcprops.JPGChannels = 4 Then
        jcprops.JPGColor = IJL_YCBCRA_FPX
        jcprops.DIBColor = IJL_RGBA_FPX
        nChannels = 4
      End If
         
      ' Create a buffer of sufficient size to hold the image:
      If cDib.Create(lJPGWidth, lJPGHeight, nChannels) Then
        ' Store DIBWidth:
        jcprops.DIBWidth = lJPGWidth
        ' Store DIBHeight:
        jcprops.DIBHeight = -lJPGHeight
        ' Store Channels:
        jcprops.DIBChannels = nChannels
            
        ' Store DIBBytes (pointer to uncompressed JPG data):
        jcprops.DIBBytes = cDib.DIBSectionBitsPtr
        ' specify align for DIB
        jcprops.DIBPadBytes = IJL_DIB_PAD_BYTES(jcprops.DIBWidth, jcprops.DIBChannels)

        Select Case jpg_scale
          Case 1
            ' Now decompress the JPG into the DIBSection:
            jerr = ijlRead(jcprops, IJL_JFILE_READWHOLEIMAGE)
          Case 2
            ' Now decompress the JPG into the DIBSection:
            jerr = ijlRead(jcprops, IJL_JFILE_READONEHALF)
          Case 4
            ' Now decompress the JPG into the DIBSection:
            jerr = ijlRead(jcprops, IJL_JFILE_READONEQUARTER)
          Case 8
            ' Now decompress the JPG into the DIBSection:
            jerr = ijlRead(jcprops, IJL_JFILE_READONEEIGHTH)
        End Select
            
            
        If jerr = IJL_OK Then
          ' convert from IJL_RGBA_FPX to BGRA
          If jcprops.DIBColor = IJL_RGBA_FPX Then
            Call ConvertFromRGBA(jcprops.DIBBytes)
          End If
          ' cDib now contains the uncompressed JPG.
          LoadJPG = True
        Else
          ' Throw error:
          Call ShowErrorMsg("FAILED TO READ IMAGE DATA " & "(" & sFile & ")", jerr)
        End If
      Else
        ' failed to create the DIB...
      End If
    End If
                        
    ' Ensure we have freed memory:
    jerr = ijlFree(jcprops)
  
  Else
    ' Throw error:
    Call ShowErrorMsg("Failed to initialise the IJL library: ", jerr)
  End If
   
End Function


Public Function SaveJPG(ByRef cDib As cDIBSection, ByVal sFile As String, qsetting As Long) As Boolean
  Dim jerr As IJLERR
  Dim jcprops As JPEG_CORE_PROPERTIES
  Dim aFile As String
  'Dim lPtr As Long
   
 jerr = ijlInit(jcprops)
 If jerr = IJL_OK Then
   ' Set up the DIB information:
   
   ' DIB width
   jcprops.DIBWidth = cDib.dib_width
   ' DIB height
   jcprops.DIBHeight = -cDib.dib_height
   ' DIB number of channels
   jcprops.DIBChannels = cDib.dib_channels
   ' DIB color space
   If jcprops.DIBChannels = 3 Then
     jcprops.DIBColor = IJL_BGR
     jcprops.JPGColor = IJL_YCBCR
     jcprops.JPGChannels = 3
     jcprops.JPGSubsampling = IJL_411
   Else
     jcprops.DIBColor = IJL_RGBA_FPX
     jcprops.JPGColor = IJL_YCBCRA_FPX
     jcprops.JPGChannels = 4
     jcprops.JPGSubsampling = IJL_4114
   End If
   ' DIBBytes (pointer to uncompressed RGB data):
   jcprops.DIBBytes = cDib.DIBSectionBitsPtr
   ' DIBPadBytes
   jcprops.DIBPadBytes = IJL_DIB_PAD_BYTES(jcprops.DIBWidth, jcprops.DIBChannels)

   ' Set up the JPEG information:
      
    aFile = StrConv(sFile, vbFromUnicode)
      
   ' JPEG filename
    jcprops.JPGFile = StrPtr(aFile)
      
   ' JPG width
   jcprops.JPGWidth = cDib.dib_width
   ' JPG height
   jcprops.JPGHeight = cDib.dib_height
   ' JPEG quality
   jcprops.jquality = qsetting '75

   ' Encode the image into file
   jerr = ijlWrite(jcprops, IJL_JFILE_WRITEWHOLEIMAGE)
   If jerr = IJL_OK Then
     SaveJPG = True
   Else
     ' Throw error
     Call ShowErrorMsg("Failed to save to JPG", jerr)
   End If
      
   ' Ensure we have freed memory:
   Call ijlFree(jcprops)
 
 Else
   ' Throw error:
   Call ShowErrorMsg("Failed to initialise the IJL library", jerr)
 
 End If
   
End Function

