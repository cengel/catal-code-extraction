Option Compare Database
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECTL
    Left As Long
    top As Long
    right As Long
    Bottom As Long
End Type

Private Type SIZEL
    cx As Long
    cy As Long
End Type
              
Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgblReterved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long 'ERGBCompression
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type


Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  'bmiColors As RGBQUAD
End Type


Private Type Bitmap
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  BmBits As Long
End Type

Private Type DIBSECTION
    dsBm As Bitmap
    dsBmih As BITMAPINFOHEADER
    dsBitfields(2) As Long
    dshSection As Long
    dsOffset As Long
End Type


' Here is the header for the Bitmap file
' as it resides in a disk file
Private Type BITMAPFILEHEADER    '14 bytes
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type



Private Declare Function apiGetObject Lib "gdi32" _
Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, _
lpObject As Any) As Long


Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long

Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEMF As Long, lpRect As RECTL) As Long

Private Declare Function apiCloseEnhMetaFile Lib "gdi32" _
Alias "CloseEnhMetaFile" (ByVal hDC As Long) As Long

'Private Declare Function apiCreateEnhMetaFile Lib "gdi32" _
'Alias "CreateEnhMetaFileA" (ByVal hDCref As Long, _
'ByVal lpFileName As String, ByVal lpRect As Any, ByVal lpDescription As String) As Long
'' lprect as RECT changed to as BYVAL as Any to allow for NULL

Private Declare Function apiCreateEnhMetaFileRECT Lib "gdi32" _
Alias "CreateEnhMetaFileA" (ByVal hDCref As Long, _
ByVal lpFileName As String, ByRef lpRect As RECTL, ByVal lpDescription As String) As Long

Private Declare Function apiDeleteEnhMetaFile Lib "gdi32" _
Alias "DeleteEnhMetaFile" (ByVal hEMF As Long) As Long

Private Declare Function GetEnhMetaFileBits Lib "gdi32" _
(ByVal hEMF As Long, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long

Private Declare Function apiSelectObject Lib "gdi32" _
 Alias "SelectObject" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hWnd As Long) As Long

Private Declare Function apiReleaseDC Lib "user32" _
  Alias "ReleaseDC" (ByVal hWnd As Long, _
  ByVal hDC As Long) As Long

Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, _
ByVal nStretchMode As Long) As Long

Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long

Private Declare Function SetViewportExtEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZEL) As Long

Private Declare Function SetViewportOrgEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long

Private Declare Function SetWindowOrgEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long

Private Declare Function SetWindowExtEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZEL) As Long

Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" _
(ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long


' CONSTANTS

' StretchBlt() Modes
Private Const BLACKONWHITE = 1
Private Const WHITEONBLACK = 2
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4
Private Const MAXSTRETCHBLTMODE = 4


' Windows Message -> ScrollBar
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
' ScrollBar Commands
Private Const SB_LINEUP = 0
Private Const SB_LINELEFT = 0
Private Const SB_LINEDOWN = 1
Private Const SB_LINERIGHT = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGELEFT = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGERIGHT = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_LEFT = 6
Private Const SB_BOTTOM = 7
Private Const SB_RIGHT = 7
Private Const SB_ENDSCROLL = 8

'  Ternary raster operations
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

' Predefined Clipboard Formats
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14

'  Mapping Modes
Private Const MM_TEXT = 1
Private Const MM_LOMETRIC = 2
Private Const MM_HIMETRIC = 3
Private Const MM_LOENGLISH = 4
Private Const MM_HIENGLISH = 5
Private Const MM_TWIPS = 6
Private Const MM_ISOTROPIC = 7
Private Const MM_ANISOTROPIC = 8

Private Const vbPicTypeNone = 0 'Picture is empty
Private Const vbPicTypeBitmap = 1 'Bitmap (.bmpBMP files)
Private Const vbPicTypeMetafile = 2 'Metafile (.wmfWMF files)
Private Const vbPicTypeIcon = 3 'Icon (.icoICO files)
Private Const vbPicTypeEMetafile = 4 'Enhanced Metafile (.emfEMF files)

' Stock Logical Objects
Private Const WHITE_BRUSH = 0
Private Const LTGRAY_BRUSH = 1
Private Const GRAY_BRUSH = 2
Private Const DKGRAY_BRUSH = 3
Private Const BLACK_BRUSH = 4
Private Const NULL_BRUSH = 5
Private Const HOLLOW_BRUSH = NULL_BRUSH
Private Const WHITE_PEN = 6
Private Const BLACK_PEN = 7
Private Const NULL_PEN = 8
Private Const OEM_FIXED_FONT = 10
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEVICE_DEFAULT_FONT = 14
Private Const DEFAULT_PALETTE = 15
Private Const SYSTEM_FIXED_FONT = 16
Private Const STOCK_LAST = 16

' Background Modes
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Const BKMODE_LAST = 2

' GetDeviceCaps
Private Const HORZSIZE = 4           '  Horizontal size in millimeters
Private Const VERTSIZE = 6           '  Vertical size in millimeters
Private Const HORZRES = 8            '  Horizontal width in pixels
Private Const VERTRES = 10           '  Vertical width in pixels
Private Const LOGPIXELSY = 90
Private Const LOGPIXELSX = 88

' How many Twips in 1 inch
Private Const TWIPSPERINCH = 1440

'-- GDI+
Private m_GDIpToken         As Long         ' Needed to close GDI+


'*******************************************
'DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 97, 2K and 2K2
'
'
'Copyright: Lebans Holdings 1999 Ltd.
'           May not be resold in whole or part by itself, or as part of a collection.
'           Please feel free to use any/all of this code within your
'           own applications, whether private or commercial, without cost or obligation.
'           Please include the one line Copyright notice if you use this function in your own code.
'
'Version:  Ver 6.0
'
'Name:      Public Function fLoadPicture(ctl As Access.Image,
'            Optional strfName As String = "", Optional AutoSize As Boolean = False) As Boolean
'Inputs:
'           ctl -> Access Image control
'           strfName -> Optional name of Image file to load and bypass File Dialog
'           Autosize -> flag to signify whether to autosize the control to its contents
'
'Purpose:   Provides functionality to load JPG,GIF,TIF,PNG,BMP,EMF,WMF,CUR and ICO
'           files on systems without the Office Graphics Filters loaded.
'           Originally developed for Systems with Access Runtime only.
'           Supports transparency in Transparent Gifs.
'           Allows you to resize Images on Forms/Reports at runtime
'           with no loss of Image quality.
'
'Dependencies: In order to have the ability to Load TIF or PNG files your system must have
'           the GDI+ Runtime DLL. With Windows XP or higher GDI+ support is native. For Windows 2K or lower
'           you must install the GDI+ DLL into the same folder as your MDB resides!!!
'           You can get the GDI+ DLL from my site or from Microsoft here:
'     Platform SDK Redistributable: GDI+ RTM
'     http://www.microsoft.com/downloads/release.asp?releaseid=32738
'
'       *******************************************************************
'       The Microsoft GDI+ DLL is freely redistributable
'       *******************************************************************
'
'Author:    Stephen Lebans
'Email:     Stephen@lebans.com
'Date:      Feb 20, 2004, 11:11:11 PM
'
'Called by: Anybody that wants to!
'
'How to use:See inline documentation within each function call.
'
'Notes:

' *************************************************************
' Version 5.0 Notes
' April 5, 2003
' Finally got around to fixing the resizing issues when control's
' SizeMode prop is set to Zoom or Stretch. With A97 any format other
' than Bitmap would resize correctly. With A2K and higher the issue was
' reversed and only Bitmap images would resize properly.
' Solution was to use Use StretchBlt with SetStretchBltMode instead
' of BitBlt in the Enhanced Metafile records.
' This current version now works in all versions of Access for
' both Forms and Reports.
' Found a silly bug/feature. Access use the Office Graphics filters
' for Bitmap and Metafile images is dependant on the letter case
' of the file extensions. I'm serious!
' Will document this and other Image handling anomolies on my Web site.
'
'

' Version 3.2 Notes
' The function was performing perfectly until I changed the Image
' Control's Size Mode property from CLIP to either ZOOM or STRETCH.
' After experimenting I've reached the following conclusions
' pertaining to a standard Image control in A97.
' 1) Any BMP or DIB file will only display properly with a Size Mode prop
' of CLIP. This statement is directed towards Bitmap files of 16 bits or
' higher. The results vary depending on the number of colors in the Image
' and how large the blocks of solid colors are. Your mileage may vary!
'
' 2) All other graphic file types will display properly with any Size Mode
' property setting.

' Why? BMP's and DIB's are an OS native file format. They are
' loaded/stored as DIB's. All other graphics formats are loaded and
' converted to DIB's and stored within a Metafile wrapper.
'
' It seems that the resizing algorithm works better on Metafiles than
' DIB's. Visually, it looks to be more of a Palette issue then the actual
' resizing routines but I haven't take the time to explore further. If
' anyone knows I would appreciate hearing from you.

' The workaround is to simply package my Bitmap within a Metafile ,
' which is the core logic employed by this function.

'Credits:
'Everyone I ever talked to in the whole world!<BG>
'
'BUGS:
'No serious bugs notices at this point in time.
'Please report any bugs to my email address.
'There is a GDI Resource leak if you play back
'several hundred Metafiles in one session. Should not showup as
'a problem in the normal course of events.
'
'What's Missing:
'Oh I'm sure there is something...there always is!
'
'HOW TO USE:
' See Example Form for how to call these functions.

'HISTORY

' Version 5.0
' Use the Render method of the StdPicture object to render the
' Image into the EMF. The load function now handles
' JPG, GIF,BMP,EMF,WMF,CUR and ICO

' Version 4.0
' Modified code that builds the EMF. Use StretchBlt with
' SetStretchBltMode to allow for smooth resizing when
' control's SizeMode prop is set to Zoom or Stretch

' Version 3.2
' Added code to simulate Magnification for the Image control.
' Autosizing of Image control to match dimensions of loaded picture.
' ScrollBars return to TOP and LEFT when CLIP is selected.

' Version 3.
' Added code to simulate ScrollBars for the Image Control.
'
' Version 2.
' Added code to have cursor change to HourGlass during the
' process to load/display the Jpeg or Gif file. Certain large
' JPEG's can take several seconds to load depending on
' the performance of the system.

' Version 1
' No stones yet!<bg>

'Enjoy
'Stephen Lebans
 '*******************************************

        


' ***************************************
' Called by the fLoadPicture function to copy the bits of the
' selected Image from a StdPicture object into our
' memory Enahnced Metafile.
' Changed to use StretchBlt and SetStretchBltMode

Function fStdPicToImageData(hStdPic As Object, ctl As Access.Image, _
Optional FileNamePath As String = "", Optional AutoSize As Boolean = False) As Boolean

' Changed all references to StdPicture to Object
' I'm going with late binding as this is a  sample database
' and many users may not be comfortable setting references.

' If you need the faster performance you can use Early binding and declare
' hStdPic as StdPicture.  Requires a Reference to Standard OLE Types.
' This file, OLEPRO32.DLL
' is usually found in your System folder. Goto the Menu Tools->References
' and set a reference to the above file.

On Error GoTo ERR_SHOWPIC

' Temp Device Context for EMF creation
Dim hDCref As Long

' DC/Window extents
Dim sz As SIZEL
Dim pt As POINTAPI
Dim rc As RECTL

' Temp var to hold API returns
Dim lngRet As Long
Dim s As String

' handle to EMF
Dim hMetafile As Long

' handle to Metafile DC
Dim hDCMeta As Long

' Array to hold binary copy of Enhanced Metafile
' we will create.
Dim arrayMeta() As Byte

' Vars to calculate resolution
Dim sngConvertX As Single
Dim sngConvertY As Single
Dim ImageWidth As Long
Dim ImageHeight As Long
Dim Xdpi As Single
Dim Ydpi As Single
Dim TwipsPerPixelX As Single
Dim TwipsPerPixely As Single
Dim sngHORZRES As Single
Dim sngVERTRES As Single
Dim sngHORZSIZE As Single
Dim sngVERTSIZE As Single


' It must be GetDC not CreateCompatibleDC!!!
hDCref = apiGetDC(0)

' Make sure user has selected a valid supported Image type
If hStdPic.Type = 0 Then
 Err.Raise vbObjectError + 523, "fStdPicToImageData.modStdPic", _
    "Sorry...This function can only read Image files." & vbCrLf & "Please Select a Valid Supported Image File"
End If

' Calculate the current Screen resolution.
' I used to simply use GetDeviceCaps and
' LOGPIXELSY/LOGPIXELSX. Unfortunately this does not yield accurate results
' with Metafiles.  LOGPIXELSY will return the value of 96dpi or 120dpi
' depending on the current Windows setting for Small Fonts or Large Fonts.
' Thanks to Feng Yuan's book "Windows Graphics Programming" for
' explaining the correct method to ascertain screen resolution.

' Let's grab the current size and resolution of our Screen DC.
sngHORZRES = apiGetDeviceCaps(hDCref, HORZRES)
sngVERTRES = apiGetDeviceCaps(hDCref, VERTRES)
sngHORZSIZE = apiGetDeviceCaps(hDCref, HORZSIZE)
sngVERTSIZE = apiGetDeviceCaps(hDCref, VERTSIZE)

' Convert millimeters to inches
sngConvertX = (sngHORZSIZE * 0.1) / 2.54
sngConvertY = (sngVERTSIZE * 0.1) / 2.54
' Convert to DPI
sngConvertX = sngHORZRES / sngConvertX
sngConvertY = sngVERTRES / sngConvertY
Xdpi = sngConvertX
Ydpi = sngConvertY


' Convert Image dimensions to Twips and Pixels
' For fun let's not convert pixels to TWIPS since
' we always do it that way. Let's be different and
' convert the StdPicture Height & Width props directly.
' These are in a Map Mode of HiMetric units, expressed in .01 mm units.

' Convert to CM
sngConvertX = hStdPic.Width * 0.001
sngConvertY = hStdPic.Height * 0.001

'Convert to Inches
sngConvertX = sngConvertX / 2.54
sngConvertY = sngConvertY / 2.54

'Convert to TWIPS
sngConvertX = sngConvertX * 1440
sngConvertY = sngConvertY * 1440
    
' Calculate TwipsPerPixel
TwipsPerPixelX = TWIPSPERINCH / Xdpi
TwipsPerPixely = TWIPSPERINCH / Ydpi

' Convert to pixels
ImageWidth = sngConvertX / TwipsPerPixelX
ImageHeight = sngConvertY / TwipsPerPixely


' Create our Enhanced Metafile - Memory Based
' Set our bounding rectangle to match that of the StdPicture object
rc.right = hStdPic.Width
rc.Bottom = hStdPic.Height

' Since this EMF may be copied to disk let's included come creator info.
s = "Stephen Lebans" & Chr(0) & Chr(0) & "www.lebans.com" & Chr(0) & Chr(0)
hDCMeta = apiCreateEnhMetaFileRECT(hDCref, vbNullString, rc, s)

' Was the EMF creation successful?
If hDCMeta = 0 Then
    Err.Raise vbObjectError + 525, "fStdPicToImageData.modStdPic", _
    "Sorry...cannot Create Enhanced Metafile"
End If

' Setup our Metafile Device Context
' Set our mapping mode
lngRet = SetMapMode(hDCMeta, MM_TEXT) 'ANISOTROPIC) 'TEXT)
' Setup the extents for our DC
lngRet = SetWindowExtEx(hDCMeta, ImageWidth, ImageHeight, sz)
lngRet = SetWindowOrgEx(hDCMeta, 0&, 0&, pt)
lngRet = SetWindowExtEx(hDCMeta, ImageWidth, ImageHeight, sz)
' Setup the basics
lngRet = SetBkMode(hDCMeta, TRANSPARENT)
lngRet = apiSelectObject(hDCMeta, GetStockObject(NULL_BRUSH))
lngRet = apiSelectObject(hDCMeta, GetStockObject(NULL_PEN))

' Fixes resizing issue.
lngRet = SetStretchBltMode(hDCMeta, COLORONCOLOR)

' Use the Render method of the StdPicture object.
' Boy the MS docs on this method are not easy to digest.
' I actually found a german web site that explained it in better detail
' with examples. My problem is that I could not find a sample
' of Rendering to a Metafile DC. The MS Docs state that you must supply
' a bounding rectangle as the last argument. No matter how I tried,
' I could not get the method to accept this parameter in any condiguration.
' Nearly ready to give up, by mistake I set the param to NULL and it worked.
' Go figure!

' The documentation regarding the Render method is not accurate.
' You do not need to suplly a bounding RECT when rendering
' to a Metafile DC.
hStdPic.Render CLng(hDCMeta), 0&, 0&, CLng(ImageWidth), CLng(ImageHeight), _
0&, hStdPic.Height, hStdPic.Width, -hStdPic.Height, vbNull

' You just never know...better safe than sorry!<grin>
DoEvents

' I have seen the following call fail on WIndos 98 systems.
' This happens when you have a report with a lot of Images
' and is due to a bug in the GDI with respect to Metafiles.
' Full details are on my web site in the Image handling FAQ.
hMetafile = apiCloseEnhMetaFile(hDCMeta)
If hMetafile = 0 Then
    fStdPicToImageData = False
    Exit Function
End If


' Grab the contents of the Metafile
lngRet = GetEnhMetaFileBits(hMetafile, 0, ByVal 0&)
If lngRet = 0 Then
    fStdPicToImageData = False
    Exit Function
End If


ReDim arrayMeta((lngRet - 1) + 8)
lngRet = GetEnhMetaFileBits(hMetafile, lngRet, arrayMeta(8))

' Delete EMF memory footprint.
lngRet = apiDeleteEnhMetaFile(hMetafile)

' If the first 40 bytes of a PictureData prop are
' not a BITMAPINFOHEADER structure then we will find
' a ClipBoard Format structure of 8 Bytes in length
' signifying whether a Metafile or Enhanced Metafile is present.
' The first 8 Bytes of a PictureData prop signify
' that the data is structured as one of the
' following ClipBoard Formats.
' CF_ENHMETAFILE
' CF_METAFILEPICT
' So the first 4 bytes tell us the format of the data.
' The next 4 bytes point to handle for a Memory Metafile.
' This is not needed for our construction purposes.
arrayMeta(0) = CF_ENHMETAFILE

' Copy our created PictureData bytes over to the Image Contol.
ctl.PictureData = arrayMeta

' Do we auto size the Image control to the
' dimensions of its the loaded image?
If AutoSize Then
    ' Error check to ensure we do not exceed
    ' SubForm boundaries
    If sngConvertX < ctl.Parent.Width Then
     '  If sngConvertX < Section(ctl.Section).Width Then
       ctl.Width = sngConvertX '+ 15
    Else
        ctl.Width = ctl.Parent.Width - 200
    End If
    
    If sngConvertY < ctl.Parent.Detail.Height Then
        ctl.Height = sngConvertY '+ 15
    Else
        ctl.Height = ctl.Parent.Detail.Height - 200
    End If
     ctl.SizeMode = acOLESizeStretch
    
End If

EXIT_SHOWPIC:
' Release our reference DC
lngRet = apiReleaseDC(0&, hDCref)
Exit Function

ERR_SHOWPIC:
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume EXIT_SHOWPIC

End Function


Public Function fLoadPicture(ctl As Access.Image, Optional strfName As String = "", Optional AutoSize As Boolean = False) As Boolean
' Inputs
' ctl -> Access Image control
' strfName -> Optional name of Image file to load and bypass File Dialog
On Error GoTo Err_fLoadPicture

' Temp Vars
Dim lngRet As Long
Dim blRet As Boolean

' Our StdPicture object returned by LoadPicture
Dim hPic As Object

' Were we passed the Optional FileName and Path
If Len(strfName & vbNullString) = 0 Then
    ' Call the File Common Dialog Window
    Dim clsDialog As Object
    Dim strTemp As String

    Set clsDialog = New clsCommonDialog

    ' Fill in our structure
    clsDialog.Filter = "All (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "JPEG (*.JPG)" & Chr$(0) & "*.JPG" & Chr$(0)
     clsDialog.Filter = clsDialog.Filter & "Tif (*.TIF)" & Chr$(0) & "*.TIF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Gif (*.GIF)" & Chr$(0) & "*.GIF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "PNG (*.PNG)" & Chr$(0) & "*.PNG" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Bitmap (*.BMP)" & Chr$(0) & "*.BMP" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Bitmap (*.DIB)" & Chr$(0) & "*.DIB" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Enhanced Metafile (*.EMF)" & Chr$(0) & "*.EMF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Windows Metafile (*.WMF)" & Chr$(0) & "*.WMF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Icon (*.ICO)" & Chr$(0) & "*.ICO" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Cursor (*.CUR)" & Chr$(0) & "*.CUR" & Chr$(0)
    
    
    clsDialog.hDC = 0
    clsDialog.MaxFileSize = 256
    clsDialog.Max = 256
    clsDialog.FileTitle = vbNullString
    clsDialog.DialogTitle = "Please Select an Image File"
    clsDialog.InitDir = vbNullString
    clsDialog.DefaultExt = vbNullString
    
    ' Display the File Dialog
    clsDialog.ShowOpen
    
    ' See if user clicked Cancel or even selected
    ' the very same file already selected
    strfName = clsDialog.FileName
    If Len(strfName & vbNullString) = 0 Then
    ' Raise the exception
      Err.Raise vbObjectError + 513, "LoadJpegGif.modStdPic", _
      "Please Select a Valid JPEG or GIF File"
    End If

' If we jumped to here then user supplied a FileName
End If

' It may take a few seconds to render larger JPEGs.
' Set the MousePointer to "HOURGLASS"
Application.Screen.MousePointer = 11
  
  
Select Case right$(strfName, 3)

    Case "bmp", "dib", "Gif", "emf", "Wmf", "ico", "cur", "jpg"
    
    ' Load the Picture as a StandardPicture object
    ' Use VBA LoadPicture function
    Set hPic = LoadPicture(strfName)
    
    Case "tif", "png"
    Dim GpInput As GdiplusStartupInput
    '-- Load the GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
      Call MsgBox("Error loading GDI+!", vbCritical)
      'Call Unload(Me)
      Exit Function
    End If
    
    ' Load the Picture as a StandardPicture object
    ' Use GDI+
    Set hPic = LoadPictureEx(strfName)
    ' Unload the GDI+ Dll
    Call mGDIpEx.GdiplusShutdown(m_GDIpToken)
    
    Case Else
    ' Unsupported Format
    ' Raise the exception
    Err.Raise vbObjectError + 518, "LoadJpegGif.modStdPic", _
    "This Image format is not supported!" & vbCrLf & strfName & vbCrLf & _
    "Please Select a Supported Image format:" & vbCrLf & _
    "JPEG, TIFF, PNG, BMP, DIB, GIF, EMF, WMF, ICO or CUR"

End Select
  
  
' Was Image loaded?
If hPic = 0 Then
    Err.Raise vbObjectError + 514, "LoadJpegGif.modStdPic", _
    "Please Select a Supported Image format:" & vbCrLf & _
    "JPEG, TIFF, PNG, BMP, DIB, GIF, EMF, WMF, ICO or CUR"
End If


' Call our function to convert the StdPicture object
' into a DIB wrapped within an Enhanced Metafile
blRet = fStdPicToImageData(hPic, ctl, , AutoSize)
' need error handling here


' Cleanup
fLoadPicture = True

Exit_LoadPic:

' Set the MousePointer back to Default
Application.Echo True
Application.Screen.MousePointer = 0
Err.Clear
Set hPic = Nothing
Set clsDialog = Nothing
Exit Function

Err_fLoadPicture:
fLoadPicture = False
Application.Screen.MousePointer = 0
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_LoadPic

End Function


' This function only works if the contents of the Image control
' are an Enhanced Metafile, which it will always be in
' this project. There is a general purpose solution
' for all formats on my site that uses the Clipboard to convert formats.
Public Function fSaveImagetoDisk(ctl As Access.Image) As Boolean

' User selected FileName from our File Dialog window
Dim sName As String
' Junk var
Dim lngRet As Long

' handle to Clipboard memory EMF
Dim hEMF As Long

' handle to our Disk based Metafile
Dim hMetafile As Long

' Array to hold binary copy of Enhanced Metafile
' we will create.
Dim arrayMeta() As Byte

' Open File Dialog
sName = fSavePicture
    If Len(sName & vbNullString) = 0 Then
    fSaveImagetoDisk = False
    Exit Function
End If

' resize our byte array to match length of PictureData prop
ReDim arrayMeta((LenB(ctl.PictureData) - 1))
arrayMeta = ctl.PictureData

' Verify CF_ENHMETAFILE
If arrayMeta(0) <> CF_ENHMETAFILE Then
    fSaveImagetoDisk = False
    MsgBox "Sorry..not a valid Enhanced Metafile contained in the Image control"
    Exit Function
End If

' Grab a local copy of the memory EMF
CopyMem hEMF, arrayMeta(4), 4
' Create a disk based copy of the Metafile
hMetafile = CopyEnhMetaFile(hEMF, sName)

' Delete EMF memory footprint.
lngRet = apiDeleteEnhMetaFile(hMetafile)


End Function



Public Function fSavePicture(Optional strfName As String = "") As String
' Inputs
' strfName -> Optional name of Image file to Save
On Error GoTo Err_fSavePicture

' Temp Vars
Dim lngRet As Long
Dim blRet As Boolean


' Were we passed the Optional FileName and Path
If Len(strfName & vbNullString) = 0 Then
    ' Call the File Common Dialog Window
    Dim clsDialog As Object
    Dim strTemp As String

    Set clsDialog = New clsCommonDialog

    ' Fill in our structure
    ' ***********************************************
    ' WARNING
    ' You must specify lowercase "emf" for the file extension. I will explain this
    ' and how it is related to the office graphics filters in detail on my web site.
       clsDialog.Filter = clsDialog.Filter & "Enhanced Metafile (*.emf)" & Chr$(0) & "*.emf" & Chr$(0)
    clsDialog.hDC = 0
    clsDialog.MaxFileSize = 256
    clsDialog.Max = 256
    clsDialog.FileTitle = vbNullString
    clsDialog.DialogTitle = "Please Enter a Valid FileName"
    clsDialog.InitDir = vbNullString
    clsDialog.DefaultExt = ".emf" 'vbNullString
    
    ' Display the File Dialog
    clsDialog.ShowSave
    
    ' See if user clicked Cancel or even selected
    ' the very same file already selected
    strfName = clsDialog.FileName
    If Len(strfName & vbNullString) = 0 Then
    ' Raise the exception
      Err.Raise vbObjectError + 513, "fSavePicture.modStdPic", _
      "Please Enter a Valid EMF Filename"
    End If

' If we jumped to here then user supplied a FileName
End If


' Cleanup
fSavePicture = strfName

Exit_SavePic:
Err.Clear
Set clsDialog = Nothing
Exit Function

Err_fSavePicture:
fSavePicture = strfName
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_SavePic

End Function


Public Sub ScrollToHome(ctl As Control)
' Scroll the Form back to X:0,Y:0
' The Form is heavily Subclassed by Access.
' It does not seem to respond to SB_TOP or SB_LEFT
' so we have to resort to the following kludge.

' Temp var
Dim lngRet As Long

' Temp counter
Dim lngTemp As Long

' Be careful because of Echo Off
On Error Resume Next

' Stop Screen Redraws
Application.Echo False

For lngTemp = 1 To 9
lngRet = SendMessage(ctl.Parent.hWnd, WM_VSCROLL, SB_PAGEUP, 0&)
lngRet = SendMessage(ctl.Parent.hWnd, WM_HSCROLL, SB_PAGELEFT, 0&)
Next lngTemp

' Start Screen Redraws
Application.Echo True

End Sub





'Public Function DIBtoPictureData(hDib As Long, ctl As Access.Image) As Boolean
'' DIBSECTION structure
'Dim ds As DIBSECTION
'' Array to hold Byte data formatted as
'' CF_DIB for the PictureData property
'Dim varTemp() As Byte
' Dim lRet As Long
'        ' Fill in our DIBSECTION structure
'        lRet = apiGetObject(hDib, Len(ds), ds)
'    If lRet = 0 Then
'    DIBtoPictureData = False
'    Exit Function
'    End If
'
'
'
'    ' Allow 40 Bytes for the DIBHeader
'    ReDim varTemp(ds.dsBmih.biSizeImage + 40)
'     If DIBnum = 0 Then
'        apiCopyMemory varTemp(40), ByVal m_lPtr, ds.dsBmih.biSizeImage
'    Else
'        apiCopyMemory varTemp(40), ByVal m_lPtr2, ds.dsBmih.biSizeImage
'    End If
'
'    apiCopyMemory varTemp(0), ds.dsBmih, 40
'
'    ' Update the PictureData property of the Image control
'     m_ImageControl.PictureData = varTemp
'    'Debug.Print "Updated PictureData Prop:" & Now
'
'End Function















' Here are some notes I attach to every
' project I do that include Metafiles and StdPicture object.
'Notes:
'1) When creating compatible DC's and Compatible Bitmaps
' make sure you use a REAL DC, not one you created!
' Had this problem before with CreateEnhancedMetafile.

'2) You cannot write directly to the Bitmap of a StdPicture
' This cost me hours to figure out. :-(
' So all you have to do is create another Memory DC and Bitmap and
' copy the StdPicture's Bitmap into that!



' Here two other methods to save the contents of an Image control to disk.
' This wonly works when the Image control contains an EMF but this is the case
' when you load any Image type other than WMF or BITMAP.

' ********************************************************************
'#1
' ********************************************************************
                 
' We are stripping off the first 8 Bytes of the
' Image1.PictureData prop and saving this to a
' disk based EMF file.

' Hold next File#
'Dim fNum As Integer
'Dim sName As String
'
'' Byte arrays to hold the PictureData prop
'Dim bArray() As Byte
'Dim cArray() As Byte
'
'' Temp var
'Dim lngRet As Long
'
'' Ensure there is data in the PcitureData prop
''If LenB(Me.JGSForm.Form.Image1.PictureData) < 108 Then Exit Sub
'
'If IsNull(Me.JGSForm.Form.Image1.PictureData) Then Exit Sub
'
'' Call the standard WIndows File Dialog
'sName = fSavePicture()
'If Len(sName & vbNullString) = 0 Then Exit Sub
'
'' Resize to hold entire PictureData prop
'ReDim bArray(LenB(Me.JGSForm.Form.Image1.PictureData) - 1)
'' Resize to hold the EMF wrapped in the PictureData prop
'ReDim cArray(LenB(Me.JGSForm.Form.Image1.PictureData) - (1 + 8))
'
'' Copy to our array
'bArray = Me.JGSForm.Form.Image1.PictureData
'
'' Copy the embedded EMF - SKIP first 8 bytes
'For lngRet = 8 To UBound(cArray) ' - (1) '+ 8)
'    cArray(lngRet - 8) = bArray(lngRet)
'Next
'
'' Get next avail file handle
'fNum = FreeFile
'
'
'' Let's Create/Open our new EMF File.
'Open sName For Binary As fNum
'
'' Write out the EMF
'Put fNum, , cArray
'
'' Close the File
'Close fNum



' ********************************************************************
'#2
' ********************************************************************

' Original Save Image control's contents to disk as an EMF
'Public Function fSaveImagetoDisk(ctl As Access.Image) As Boolean
'Dim sName As String
'Dim lngRet As Long
'Dim hEMF As Long
'' handle to Clipboard memory EMF
'Dim hMetafile As Long
'
'' handle to Metafile DC
''Dim hDCMeta As Long
'
'' Array to hold binary copy of Enhanced Metafile
'' we will create.
'Dim arrayMeta() As Byte
'
'' Open File Dialog
'sName = fSavePicture
'    If Len(sName & vbNullString) = 0 Then
'    fSaveImagetoDisk = False
'    Exit Function
'End If
'
'lngRet = FPictureDataToClipBoard(ctl)
'' Geta handle to an EMF from the Clipboard
'hMetafile = GetClipBoard(CF_ENHMETAFILE)
'
'' Grab the contents of the Metafile
'lngRet = GetEnhMetaFileBits(hMetafile, 0, ByVal 0&)
'ReDim arrayMeta((lngRet - 1) + 8)
'lngRet = GetEnhMetaFileBits(hMetafile, lngRet, arrayMeta(8))
'
'' Delete EMF memory footprint.
'lngRet = apiDeleteEnhMetaFile(hMetafile)
'
'' If the first 40 bytes of a PictureData prop are
'' not a BITMAPINFOHEADER structure then we will find
'' a ClipBoard Format structure of 8 Bytes in length
'' signifying whether a Metafile or Enhanced Metafile is present.
'' The first 8 Bytes of a PictureData prop signify
'' that the data is structured as one of the
'' following ClipBoard Formats.
'' CF_ENHMETAFILE
'' CF_METAFILEPICT
'' So the first 4 bytes tell us the format of the data.
'' The next 4 bytes point to handle for a Memory Metafile.
'' This is not needed for our construction purposes.
'arrayMeta(0) = CF_ENHMETAFILE
'
'
'' Save to disk
'Dim fNum As Integer
'
'
'' Byte arrays to hold the PictureData prop
'Dim bArray() As Byte
'Dim cArray() As Byte
'
'
'
'' Resize to hold entire PictureData prop
''ReDim bArray(LenB(Me.JGSForm.Form.Image1.PictureData) - 1)
'' Resize to hold the EMF wrapped in the PictureData prop
'ReDim cArray(UBound(arrayMeta) - (8))
'
'' Copy to our array
''bArray = Me.JGSForm.Form.Image1.PictureData
'
'' Copy the embedded EMF - SKIP first 8 bytes
'For lngRet = 8 To UBound(cArray) ' - (1) '+ 8)
'    cArray(lngRet - 8) = arrayMeta(lngRet)
'Next
'
'' Get next avail file handle
'fNum = FreeFile
'
'
'' Let's Create/Open our new EMF File.
'Open sName For Binary As fNum
'
'' Write out the EMF
'Put fNum, , cArray
'
'' Close the File
'Close fNum
'
'
'
'
'End Function
'

