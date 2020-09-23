Attribute VB_Name = "mAlphaBlt"
Option Explicit
' I have made modifications to this module and annotated my changes in the routines.

' The mods made here are a result of testing the routines against my PNG project that
'   I thought I'd post this month, but will wait for another week or 2 or 3 to finalize.
' Anyway, Carles' routines are most excellent. For my PNG project, I wanted full
' stretch capabilities, including stretching a portion of an image vs the entire image.
' I have found his resizing routines to out perform (quality-wise) StretchDIBits &
' StretchBlt; therefore, I plan on using it in my PNG project. However, the resizing
' routines did not allow portion stretching & calculated unecessary pixel blends for
' images that extended beyond the physical boundaries of the destination DC/bitmap.

' Unmodified module(s) found at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60424&lngWId=1
' Primary changes made for this sample project:
'   - Removed interpolation routine, not needed for my project
'   - Combined the AlphaBlend/AlphBlendStretch into a single module
'   - Removed the 2 AlphaBlt modules
'   - Added capability of providing a global alpha parameter
'   - Modified/rewrote code to use clipped target and source areas. Explained below
'   - Added 2 new local routines to calculate the clipped areas
'   - Added support code to alphablend non-32 bit images
'   - Made minor speed tweaks where possible. Tough to improve; so well-written
'   - Modified again to clip on negative destination DC offsets & to fix minor errors
'   - Modified again to add global alpha processing in its own loop. A majority of
'       the time, I would imagine global alpha is not used. Calculating the total
'       alpha to include global alpha requires 2 extra math statements for each pixel.
'       By not performing those statements, we recover signifcant speed lost when
'       global alphas are not used. Modified, minimally, a few routines for flow.

' ABOUT CLIPPING AND STRETCHING
' The original code, excellent as it is, didn't allow the user to stretch only
' a portion of the bitmap. Also when stretching the entire bitmap, it was possible
' to waste signfiicant time calculating pixels that would never by used....

' A StretchBlt type function whether AlphaBlend, StretchBlt, StretchDIBits, etc
' has 2 sets of boundaries: one for the target DC and the other for the source
' DC/image. The orignal routines didn't offer the second set. To accomplish this
' the routines need to calculate/map pixel locations between the areas of the
' target and source; and these areas may be different in size, position, and scale.

' Now when API stretching an image where the size would extend past the boundaries of
' the target DC, obviously those pixels won't be displayed; the API clips them.
' So, why should we process those pixels? By calculating clipping, we can prevent
' pixels never used from being processed. However, this adds some overhead to
' the routines and slightly slows down the processing of an image that is
' completely contained within the target DC. On the other hand, if an image is
' stretched beyond the DC's physcal boundaries, the speed gained can be significant.
' When reduced/stretched to a size contained by the DC, no speed is gained.
' The clipped area will never be greater than the stretched/actual (S/A) area.

' Cost Savings Formula is (S/A Ht * (S/A Wd *4)) - (Clipped Ht * (Clipped Wd *4))
'   Using the penguin image (735x783) on a 500x500 DC as an example...
'       (783*(735*4))-(500*(500*4)) = 1,302,020 less bytes processed

' The overhead mentioned above is required because we need to track 8 different
' measurements (source:x,y,w,h & destination:x,y,w,h) across 2 different scales/ratios,
' synchronize/shift the different areas over a 3rd area (temporary DIB used for the
' blending) and then apply all these offsets within the pvResize & AlphablendStretch
' functions. By far, the most difficult contribution to Carles' routines
'===================================================================================
' original module header follows.


'================================================
' Module:        mAlphaBlt.bas
' Author:        Carles P.V.
'                (see resizing routines credits)
' Dependencies:
' Last revision: 2005.05.08
'================================================
'
' History:
'
' - 2005.04.01: First release
'
' - 2005.05.03: Speed up: checked special alpha values
'               (full opaque and full transparent)
'
' - 2005.05.08: AlphaBltStretch and AlphaBlendStretch variations:
'
'               - 'Bilinear resize' original routine from 'Reconstructor' by Peter Scale
'                 (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46515&lngWId=1)
'
'               - 'Integer maths' version from 'RVTVBIMG v2 - Image Processing in VB' by Ron van Tilburg
'                 (http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=47445&lngWId=1)
'
'               Slight modifications:
'
'               - Reduced to 32-bit case.
'
'               - X and Y axes scaling LUTs.
'
'               If someone wants to use these AlphaBlendXXX (interpolated) functions in a "multi-layer"
'               application, will finish up with undesired results. These results can be appreciated if
'               you set iterations at, for example, 10. We finish up with really darken edge-pixels.
'               A correct interpolated resizing of alpha-bitmaps is quite more complex.
'               Problems come from edge blended pixels (pre-blending (interpolation) null alpha pixels
'               color information...)


Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type SafeArrayBound
    cElements As Long
    lLbound   As Long
End Type

Private Type SafeArray1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds     As SafeArrayBound
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7
Private Const COLORONCOLOR = 3

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (ptr() As Any) As Long

'// added by LaVolpe
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'//

Public Function AlphaBlendStretch( _
                ByVal hDC As Long, _
                ByVal dstX As Long, ByVal dstY As Long, _
                ByVal dstWidth As Long, ByVal dstHeight As Long, _
                ByVal hBitmap As Long, _
                ByVal srcX As Long, ByVal srcY As Long, _
                Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                Optional ByVal GlobalAlpha As Byte = 255) As Long

'/modified by LaVolpe:
' added src_  parameters & GlobalAlpha parameter above, removed Interpolate parameter
'-- check for quick aborts - invalid parameters passed
  If GlobalAlpha = 0 Then Exit Function ' image will be completely transparent
  If srcWidth < 0 Then Exit Function    ' optional but not allowed to be negative
  If srcHeight < 0 Then Exit Function   ' optional but not allowed to be negative
  If dstWidth < 1 Or srcX < 0 Then Exit Function
  If dstHeight < 1 Or srcY < 0 Then Exit Function
'//

  Dim uBI       As BITMAP
  Dim uBIH      As BITMAPINFOHEADER

  Dim lhDC      As Long
  Dim lhDIB     As Long
  Dim lhDIBOld  As Long

  Dim a1        As Long
  Dim a2        As Long

  Dim uSSA      As SafeArray1D
  Dim aSBits()  As Byte
  Dim aSBitsR() As Byte
  Dim uDSA      As SafeArray1D
  Dim aDBits()  As Byte
  Dim lpData    As Long

'// Added to allow use of clipping calculations & global alpha setting
  Dim srcClip As RECT   ' modified dstX,dstY,dstWidth,dstHeight values
  Dim dstClip As RECT   ' modified srcX,srcY,srcWidth,srcHeight values
  Dim dstBI As BITMAP
  Dim srcScanWidth As Long, dstScanWidth As Long
  Dim srcPos As Long, dstPos As Long
  Dim srcRow As Long, pixelLoc As Long
  Dim dstRow As Long
  Dim h_Non32bit As Long  ' handle for non 32bit image to be blended
  Dim gAlpha As Long
  
  Const BytesPerPixel As Long = 4&
'//

  '-- Check type (bitmap)

    If (GetObjectType(hBitmap) = OBJ_BITMAP) Then

        '-- Get bitmap info
        If (GetObject(hBitmap, Len(uBI), uBI)) Then

            '-- Check if source bitmap is 32-bit!
            '// modified by LaVolpe: support non 32bit and 32bit non-dibs
            ' Trial & Error shows that sometimes 32bit loaded in stdPicture may not
            ' contain the needed .bmBits pointer; therefore, convert to DIB
            If (uBI.bmBits = 0) Or (uBI.bmBitsPixel <> 32) Then
                ' slightly speeds up processing by using the global alpha for the
                ' converted non-32bit alpha values
                h_Non32bit = lvConvertTo32Bit(hBitmap, uBI, GlobalAlpha)
                If h_Non32bit = 0 Then Exit Function
            End If

            '// added by LaVolpe. Attempt to get bitmap size in passed DC
            GetObject GetCurrentObject(hDC, OBJ_BITMAP), Len(dstBI), dstBI
            
            '   Clip the source & destination areas to include only needed pixels
            SetRect dstClip, dstX, dstY, dstWidth, dstHeight
            If lvClipDestination(dstClip, dstBI.bmWidth, dstBI.bmHeight) = False Then Exit Function
            ' following function validates optional srcWidth & srcHeight parameters also
            SetRect srcClip, srcX, srcY, 0, 0   ' last 2 will be filled in by next function
            If lvClipSource(srcClip, dstClip, uBI.bmWidth, uBI.bmHeight, _
                            srcWidth, srcHeight, dstWidth, dstHeight) = False Then Exit Function
            ' any dstX/Y non-zero offsets were handled in above two functions.
            ' The processed/blended image needs to be shifted towards 0,0 as needed
            If dstX < 0 Then dstX = 0
            If dstY < 0 Then dstY = 0
                
            dstScanWidth = BytesPerPixel * dstClip.Right
            '//^ scan width of the destination area, may be different than the source
        
            With uBIH

                '-- Define DIB info
                .biSize = Len(uBIH)
                .biPlanes = 1
                .biBitCount = 32
                .biWidth = dstClip.Right
                .biHeight = dstClip.Bottom
                .biSizeImage = dstScanWidth * .biHeight
                
            End With

            '-- Create a temporary DIB section, select into a DC, and
            '   bitblt destination DC area
            '/ Modified by LaVolpe: added safety check should something go wrong
            lhDC = CreateCompatibleDC(0)
            If lhDC = 0 Then Exit Function
            lhDIB = CreateDIBSection(lhDC, uBIH, DIB_RGB_COLORS, lpData, 0, 0)
            If lhDIB = 0 Then
                Call DeleteDC(lhDC)
                Exit Function
            End If
            lhDIBOld = SelectObject(lhDC, lhDIB)
            
            Call BitBlt(lhDC, 0, 0, dstClip.Right, dstClip.Bottom, hDC, dstX, dstY, vbSrcCopy)
            Call SelectObject(lhDC, lhDIBOld)
            Call DeleteDC(lhDC)

            On Error GoTo ExitRoutine
            '-- Map destination color data
            Call pvMapDIBits(uDSA, aDBits(), lpData, uBIH.biSizeImage)

            '-- Map source color data
            Call pvMapDIBits(uSSA, aSBits(), uBI.bmBits, uBI.bmWidthBytes * uBI.bmHeight)
            
            '-- Resize source color data
            '//modified by LaVolpe.
            If (dstWidth <> srcWidth Or dstHeight <> srcHeight) Then
                aSBitsR() = pvResize(aSBits(), uBI.bmWidth, uBI.bmHeight, _
                                    dstWidth, dstHeight, dstClip, srcClip)
                srcScanWidth = dstScanWidth
                ' the previous function aligns aSBitsR() & aDBits() so they are equal
                ' in size (different scales). So the bytes per scan are the same
            Else
                aSBitsR() = aSBits()
                srcScanWidth = uBI.bmWidthBytes ' original source bytes per scan
            End If
            '//
                
            '-- Blend with destination
            
            '// Modified by LaVolpe. Use global alpha blend and clipping in calculations
            With srcClip
                
                If GlobalAlpha = 255 Then
                    
                    ' put in own loop to prevent doing 2 extra calculations per pixel
                    ' This speeds up non-globally blended images
                    For srcRow = .Bottom To .Top
                        pixelLoc = srcRow * srcScanWidth + .Left
                        For srcPos = pixelLoc To pixelLoc + .Right - &H1 Step BytesPerPixel
                        
                            a1 = aSBitsR(srcPos + &H3)
                            If (a1 = &HFF) Then
                                '-- Dest. = Source
                                CopyMemory aDBits(dstPos), aSBitsR(srcPos), &H3
                                
                            ElseIf (a1 > &H0) Then
                                '-- Blend
                                a2 = &HFF - a1
                                
                                aDBits(dstPos) = (a1 * aSBitsR(srcPos) + a2 * aDBits(dstPos)) \ &HFF
                                aDBits(dstPos + &H1) = (a1 * aSBitsR(srcPos + &H1) + a2 * aDBits(dstPos + &H1)) \ &HFF
                                aDBits(dstPos + &H2) = (a1 * aSBitsR(srcPos + &H2) + a2 * aDBits(dstPos + &H2)) \ &HFF
                           
                           ' Else
                                '-- Do nothing (dest. preserved)
                            End If
                            dstPos = dstPos + BytesPerPixel
                        Next srcPos
                        dstRow = dstRow + &H1
                        dstPos = dstRow * dstScanWidth
                    Next srcRow
                    
                Else ' using global alpha setting
                
                    gAlpha = &H64& * GlobalAlpha ' x100 to use in integer math
                    For srcRow = .Bottom To .Top
                        pixelLoc = srcRow * srcScanWidth + .Left
                        For srcPos = pixelLoc To pixelLoc + .Right - &H1 Step BytesPerPixel
                        
                            a1 = (aSBitsR(srcPos + &H3) * gAlpha)
                            '^ use global alpha setting in calculations
                        
                            If (a1 = &H633864) Then  ' &H633864=255*25500
                                '-- Dest. = Source
                                '/modified by LaVolpe: replaced 3 byte assignments with CopyMemory
                                CopyMemory aDBits(dstPos), aSBitsR(srcPos), &H3
                                
                            ElseIf (a1 > &H0) Then
                                '-- Blend
                                a1 = a1 \ &H639C    ' 25500&
                                a2 = &HFF - a1
                                
                                aDBits(dstPos) = (a1 * aSBitsR(srcPos) + a2 * aDBits(dstPos)) \ &HFF
                                aDBits(dstPos + &H1) = (a1 * aSBitsR(srcPos + &H1) + a2 * aDBits(dstPos + &H1)) \ &HFF
                                aDBits(dstPos + &H2) = (a1 * aSBitsR(srcPos + &H2) + a2 * aDBits(dstPos + &H2)) \ &HFF
                           
                           ' Else
                                '-- Do nothing (dest. preserved)
                            End If
                            dstPos = dstPos + BytesPerPixel
                        Next srcPos
                        dstRow = dstRow + &H1
                        dstPos = dstRow * dstScanWidth
                    Next srcRow
                End If
                
            End With

            '-- Paint alpha-blended (stretched)
            AlphaBlendStretch = StretchDIBits(hDC, dstX, dstY, dstClip.Right, dstClip.Bottom, 0, 0, dstClip.Right, dstClip.Bottom, ByVal lpData, uBIH, DIB_RGB_COLORS, vbSrcCopy)
                
ExitRoutine: '-- Unmap
            Call pvUnmapDIBits(aDBits())
            Call pvUnmapDIBits(aSBits())

            '-- Clean up
            Call DeleteObject(lhDIB)
            ' if non-32bit dib converted to 32bit then destroy it
            If h_Non32bit <> 0 Then DeleteObject h_Non32bit
        
        End If
    End If
End Function

Public Function pvResize(ByRef aOldBits() As Byte, _
                         ByVal OldWidth As Long, ByVal OldHeight As Long, _
                         ByVal NewWidth As Long, ByVal NewHeight As Long, _
                         ByRef dstSize As RECT, ByRef srcSize As RECT _
                         ) As Byte()
                            
'Note: Slight difference when resizing (nearest) via GDI and via native routine ('rounding' issues).
'      This can be observed when scale factor is not integer.
  
'// modified by LaVolpe. A majority of this routine was tweaked to also
'   allow stretching a portion of the image. Additionally, offsets/clipping is
'   employed to dramatically speed processing any image that is stretched beyond
'   the physical boundaries of the target DC/bitmap.

'  This is where all the clipping offsets come into play. We are mapping pixels
'  from one area (Source Image) to another area (Destination DC/Image) keeping
'  track of the different lefts, tops, widths, heights, and scales. This routine
'  was well written enough that I basically only had to insert those offsets
'  in the appropriate calculations.
  
    Dim aNewBits() As Byte
    
    Dim po As Long, qo As Long
    Dim xn As Long, yn As Long, qn As Long
    
    Dim OldScan         As Long
    Dim NewScan         As Long
    Dim BytesPerPixel   As Long
    Dim xLU()           As Long
  
    ' Scan lines / pixel width
    BytesPerPixel = 4
    OldScan = BytesPerPixel * OldWidth
    NewScan = BytesPerPixel * dstSize.Right
    
    ' Resized 'bits' array
    ReDim aNewBits(0 To NewScan * dstSize.Bottom - 1)
    
    ' Scaled fractions
    '/ modified by LaVolpe. Replaced Double with Longs\Integer division
    Dim kX As Long, kY As Long
    
    With srcSize
    
        kX = (.Right * 100) \ NewWidth
        kY = (.Bottom * 100) \ NewHeight
    
        ' Scaling LUTs
        ReDim xLU(0 To dstSize.Right - 1) As Long
        
        '/ modified by LaVolpe. Allows negative offsets & portion stretching
        '... calculate relative column position in relation to scaled version
        For xn = 0 To dstSize.Right - 1
            xLU(xn) = ((((xn + dstSize.Left) * kX) \ 100) + .Left) * BytesPerPixel
        Next xn
        
        .Bottom = OldHeight - .Top - 1
        '^ for bottom up dibs, set position of top row in relation to
        '   user-passed srcY in the AlphaBlendStretch function
        
        For yn = 0 To dstSize.Bottom - 1
            po = (.Bottom - (((yn + dstSize.Top) * kY) \ 100)) * OldScan
            '^ position of source row relative to scaled version, allowing negative offsets
            qn = (dstSize.Bottom - 1 - yn) * NewScan
            '^ current scanline for the scaled image
            For xn = 0 To dstSize.Right - 1
                ' nearest relative raw pixel
                CopyMemory aNewBits(qn), aOldBits(po + xLU(xn)), BytesPerPixel
                qn = qn + BytesPerPixel
            Next xn
        Next yn
        
    End With
    
    SetRect srcSize, 0, dstSize.Bottom - 1, dstSize.Right * 4, 0
    pvResize = aNewBits

End Function

Private Sub pvMapDIBits(uSA As SafeArray1D, aBits() As Byte, ByVal lpData As Long, ByVal lSize As Long)
    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = lSize
        .pvData = lpData
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIBits(aBits() As Byte)
    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub

Private Function lvClipDestination(ByRef TargetR As RECT, _
                            ByVal Width As Long, ByVal Height As Long) As Boolean

' Used to restrict alphablending & stretching to the physical boundaries of the
' target DC/bitmap.

' When AutoRedaw is True or an off-screen bitmap is used, the DC contains
' a bitmap that has boundaries. Drawing over those boundaries is useless since the
' pixels won't be added to that target area anyway, so why do the calculations &
' pixel modifications for pixels that won't be kept? We won't.

' Note about AutoRedraw on Forms/PictureBoxes. If the object is scalable,
' the bitmap size for that object is screen width x screen height.
' If the object is not scalable and AutoRedraw is True, then the
' bitmap is the same size as the object. In these cases and the case with a
' memory/offscreen DC, the bitmap has a set size & we can avoid processing
' any pixels that do not "fall" on that bitmap.

    ' When function exits. the TargetR structure will contain:
    '   Left = ABS(negative X offset), if any, else zero. Used by pvResize
    '   Top =  ABS(negative Y offset), if any, else zero. Used by pvResize
    '   Right = the actual, minimal, width needed to process this image
    '   Bottom = the actual, minimal, height needed to process this image
    
    If TargetR.Right < 0 Or TargetR.Bottom < 0 Then Exit Function

    ' Height/Width will be zero if target DC is a VB DC with autoRedraw=False
    If Height = 0 Then Height = Screen.Height \ Screen.TwipsPerPixelY
    If Width = 0 Then Width = Screen.Width \ Screen.TwipsPerPixelX
    
    With TargetR
        If .Left >= 0 Then
            ' subtract offset from available width if possible
            If .Left + .Right > Width Then .Right = Width - .Left
            .Left = 0
        Else    ' negative offsets
            .Right = .Right + .Left ' subtract offset from available width
            If .Right > Width Then .Right = Width
            .Left = -.Left          ' offset will be applied to the source also
        End If
        
        
        ' do the same for the vertical boundaries
        If .Top >= 0 Then
            ' subtract offset from available height if possible
            If .Top + .Bottom > Height Then .Bottom = Height - .Top
            .Top = 0
        Else    ' negative offsets
            .Bottom = .Bottom + .Top ' subtract offset from available height
            If .Bottom > Height Then .Bottom = Height
            .Top = -.Top             ' offset will be applied to the source also
        End If
        
        ' if user passed an invalid range of values, we don't process anything
        lvClipDestination = (.Right > 0 And .Bottom > 0)
        
    End With
    
End Function

Private Function lvClipSource(ByRef SourceR As RECT, ByRef TargetR As RECT, _
                         ByVal Width As Long, ByVal Height As Long, _
                         ByRef imgWidth As Long, imgHeight As Long, _
                         ByVal drawWidth As Long, ByVal drawHeight As Long) As Boolean

    ' If image is being stretched, then the pvResize function modifies it instead;
    '   otherwise, when function exits, the SourceR structure will contain:
    '   Left = any left offset into the source image multiplied by 4
    '   Right = the number of pixels per scan line to be processed multiplied by 4
    '   Top = the last scan line of the image to be processed
    '   Bottom = the 1st scan line of the image to be processed

    ' ensure these passed values are validated.
    ' These were optional when sent to AlphaBlendStretch function
    If imgWidth = 0 Or imgWidth > Width Then imgWidth = Width
    If imgHeight = 0 Or imgHeight > Height Then imgHeight = Height

    With SourceR
    
        ' tweak any passed boundaries to ensure they are within the pysical
        ' boundaries of the image to be drawn. AutoFix the passed
        ' imgWidth & imgHeight parameters if greater than actual image size
        
        If .Left + imgWidth > Width Then imgWidth = Width - .Left
        If .Top + imgHeight > Height Then imgHeight = Height - .Top
        .Right = imgWidth
        .Bottom = imgHeight
        
        If drawWidth = imgWidth And drawHeight = imgHeight Then
        
            ' when not stretching the image, we want the smallest of the
            ' target area and image area. This is what will be processed.
            
            If TargetR.Right < .Right Then .Right = TargetR.Right _
                Else TargetR.Right = .Right
            
            If TargetR.Bottom < .Bottom Then .Bottom = TargetR.Bottom _
                Else TargetR.Bottom = .Bottom
            
            If (.Right > 0 And .Bottom > 0) Then
            ' if user passed an invalid range of values, we don't process anything
            
                ' Set up the structure used in the alphablending routine.
                ' When TargetR comes to this routine, if negative offsets were
                ' used for the position on the target DC, then those values (.Left/.Top)
                ' were converted to positive in the lvClipDestination routine else they
                ' were changed to zeros.
                
                .Bottom = Height - (.Top + TargetR.Top + .Bottom)
                .Top = .Bottom + TargetR.Bottom - 1
                .Right = .Right * 4
                .Left = (.Left + TargetR.Left) * 4
                
                lvClipSource = True
            
            End If
            
        Else
            
            ' if user passed an invalid range of values, we don't process anything
            lvClipSource = (.Right > 0 And .Bottom > 0)
        
        End If
        
    End With
    
    
End Function


Private Function lvConvertTo32Bit(ByVal h_Source As Long, ByRef sourceBI As BITMAP, ByRef gAlpha As Byte) As Long

    ' Pretty straightforward. Create 32bit image & BitBlt non-32bit image over it

    Dim lhDC    As Long
    Dim lhDIB   As Long
    Dim hDIBold As Long
    
    Dim thDC    As Long
    Dim tOldBmp As Long
    Dim uBIH    As BITMAPINFOHEADER
    
    Dim uSSA      As SafeArray1D
    Dim aSBits()  As Byte
    Dim I         As Long

    With uBIH
        .biBitCount = 32
        .biHeight = sourceBI.bmHeight
        .biWidth = sourceBI.bmWidth
        .biPlanes = 1
        .biSize = Len(uBIH)
    End With
        
    sourceBI.bmWidthBytes = 4 * sourceBI.bmWidth    ' element used in main routine

    lhDC = CreateCompatibleDC(0)
    If lhDC <> 0 Then
        lhDIB = CreateDIBSection(lhDC, uBIH, DIB_RGB_COLORS, sourceBI.bmBits, 0, 0)
        
        If lhDIB <> 0 Then
            thDC = CreateCompatibleDC(0)
            
            If thDC = 0 Then
            
                ' can't blt non-24bit to DIB (memory issues with pc)
                DeleteObject lhDIB
                lhDIB = 0
                
            Else
            
                hDIBold = SelectObject(lhDC, lhDIB)
                tOldBmp = SelectObject(thDC, h_Source)
                Call BitBlt(lhDC, 0, 0, sourceBI.bmWidth, sourceBI.bmHeight, thDC, 0, 0, vbSrcCopy)
            
                ' clean up
                Call SelectObject(lhDC, hDIBold)
                Call SelectObject(thDC, tOldBmp)
                Call DeleteDC(thDC)
                
                ' now if the original bits were not 32bit, then we need to update the alphas
                If sourceBI.bmBitsPixel <> 32 Then
                
                    Call pvMapDIBits(uSSA, aSBits(), sourceBI.bmBits, sourceBI.bmWidthBytes * sourceBI.bmHeight)
                    For I = 3 To UBound(aSBits) Step 4
                        aSBits(I) = gAlpha ' use the passed global alpha value
                    Next
                    Call pvUnmapDIBits(aSBits)
                    gAlpha = 255    ' by setting to 255, the calling routine's faster
                                    ' loop will be triggered for global alphablending
                End If
                    
            End If
        
        End If
        Call DeleteDC(lhDC)
        
    End If
    
    lvConvertTo32Bit = lhDIB
End Function
