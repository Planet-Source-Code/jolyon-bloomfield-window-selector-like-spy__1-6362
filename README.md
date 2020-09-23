<div align="center">

## Window Selector \(Like SPY\+\+\)


</div>

### Description

Here's a little bit of code that I'm using in one of my apps. I slaved over it for 3 nights, so I thought that others might also like to use it. Basically, it Displays a bounding box around the window that the cursor is over, just like selecting a window in Microsoft SPY++. (I decompiled it and monitored it and used dependancies and all sorts of junk from there too, in hope of imitating it). So here it is.

This code snippet uses window regions, hDC's, System Objects, and basically complex drawing API's. It also shows how to clean up after yourself (I.e., Release memory correctly).

Although it uses advanced techniques, you should get an idea of how graphics sort of work from here. Highly commented code, so as to maximise the knowledge that you can get from here.

Please leave comments, as the feedback makes the program. If you like it, please vote :o)

Jolyon Bloomfield February 2000
 
### More Info
 
Fancy Stuph using lots of Window API calls

None known, unless your app hangs


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jolyon Bloomfield](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jolyon-bloomfield.md)
**Level**          |Advanced
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jolyon-bloomfield-window-selector-like-spy__1-6362/archive/master.zip)





### Source Code

```
'
' Sorry I haven't put the declarations in it's box and all that jazz, but I did it like this
' So that you could select it all and just place it into a new form.
' Yes, it's that easy. Create a new form, copy and paste this code.
' Then click on the form and hold down the mouse, and drag over to another window.
'
' Jolyon Bloomfield, February 2000
'
' A note to using this code: I guess since I've put it here, anybody can use it.
' If you do, please give me credit for the hard work that I put into this.
' It wasn't an easy process, and I don't want anybody taking credit for my work.
'
'
' The only bug I've found, is that when a window is maximised, it has coordinates
' that exceed the bounding area of the screen. I tried to offset this effect,
' but gave up.
'
Option Explicit      ' Require variable Declaration
' PointAPI and RECT are the two most common structures used in graphics in Windows
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long  ' Get the cursor position
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long  ' Get the handle of the window that is foremost on a particular X, Y position. Used here to get the window under the cursor
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long   ' Get the window co-ordinates in a RECT structure
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long   ' Retrieve a handle for the hDC of a window
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long   ' Release the memory occupied by an hDC
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long  ' Create a GDI graphics pen object
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long  ' Used to select brushes, pens, and clipping regions
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long   ' Get hold of a "stock" object. I use it to get a Null Brush
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long  ' Used to set the Raster OPeration of a window
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long  ' Delete a GDI Object
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long  ' GDI Graphics- draw a rectangle using current pen, brush, etc.
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long   ' Set mouse events only for one window
Private Declare Function ReleaseCapture Lib "user32" () As Long    ' Release the mouse capture
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long   ' Create a rectangular region
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long   ' Select the clipping region of an hDC
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long    ' Get the Clipping region of an hDC
Private Const NULL_BRUSH = 5  ' Stock Object
Private Selecting As Boolean   ' Amd I currently selecting a window?
Private BorderDrawn As Boolean    ' Is there a border currently drawn that needs to be undrawn?
Private Myhwnd As Long     ' The current hWnd that has a border drawn on it
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Set the selecting flag
Selecting = True
' Capture all mouse events to this window (form)
SetCapture Me.hwnd
' Simulate a mouse movement event to draw the border when the mouse button goes down
Form_MouseMove 0, Shift, X, Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Security catch to make sure that the graphics don't get mucked up when not selecting
If Selecting = False Then Exit Sub
' Call the "Draw" subroutine
Draw
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If not selecting, then skip
If Selecting = False Then Exit Sub
' Clean up the graphics drawn
UnDraw
' Release mouse capture
ReleaseCapture
' Not selecting
Selecting = False
' Reset the variable
Myhwnd = 0
End Sub
Private Sub Draw()
Dim Cursor As POINTAPI    ' Cursor position
Dim RetVal As Long      ' Dummy returnvalue
Dim hdc As Long        ' hDC that we're going to be using
Dim Pen As Long        ' Handle to a GDI Pen object
Dim Brush As Long       ' Handle to a GDI Brush object
Dim OldPen As Long      ' Handle to previous Pen object (to restore it)
Dim OldBrush As Long     ' Handle to previous brush object (to restore it)
Dim OldROP As Long      ' Value of the previous ROP
Dim Region As Long      ' Handle to a GDI Region object that I create
Dim OldRegion As Long     ' Handle to previous Region object for the hDC
Dim FullWind As RECT     ' the bounding rectangle of the window in screen coords
Dim Draw As RECT       ' The drawing rectangle
'
' Getting all of the ingredients ready
'
' Get the cursor
GetCursorPos Cursor
' Get the window
RetVal = WindowFromPoint(Cursor.X, Cursor.Y)
' If the new hWnd is the same as the old one, skip drawing it, so to avoid flicker
If RetVal = Myhwnd Then Exit Sub
' New hWnd. If there is currently a border drawn, undraw it.
If BorderDrawn = True Then UnDraw
' Set the BorderDrawn property to true, as we're just about to draw it.
BorderDrawn = True
' And set the hWnd to the new value.
' Note, I didn't do it before, because the UnDraw routine uses the Myhwnd variable
Myhwnd = RetVal
' Print the hWnd on the form in Hex (just so see what windows are at work)
Me.Cls
Me.Print Hex(Myhwnd)
' You could extract other information from the window, such as window title,
' class name, parent, etc., and print it here, too.
' Get the full Rect of the window in screen co-ords
GetWindowRect Myhwnd, FullWind
' Create a region with width and height of the window
Region = CreateRectRgn(0, 0, FullWind.Right - FullWind.Left, FullWind.Bottom - FullWind.Top)
' Create an hDC for the hWnd
' Note: GetDC retrieves the CLIENT AREA hDC. We want the WHOLE WINDOW, including Non-Client
' stuff like title bar, menu, border, etc.
hdc = GetWindowDC(Myhwnd)
' Save the old region
RetVal = GetClipRgn(hdc, OldRegion)
' Retval = 0: no region   1: Region copied  -1: error
' Select the new region
RetVal = SelectObject(hdc, Region)
' Create a pen
Pen = CreatePen(DrawStyleConstants.vbSolid, 6, 0)  ' Draw Solid lines, width 6, and color black
' Select the pen
' A pen draws the lines
OldPen = SelectObject(hdc, Pen)
' Create a brush
' A brush is the filling for a shape
' I need to set it to a null brush so that it doesn't edit anything
Brush = GetStockObject(NULL_BRUSH)
' Select the brush
OldBrush = SelectObject(hdc, Brush)
' Select the ROP
OldROP = SetROP2(hdc, DrawModeConstants.vbInvert)  ' vbInvert means, whatever is draw,
     ' invert those pixels. This means that I can undraw it by doing the same.
'
' The Drawing Bits
'
' Put a box around the outside of the window, using the current hDC.
' These coords are in device co-ordinates, i.e., of the hDC.
With Draw
 .Left = 0
 .Top = 0
 .Bottom = FullWind.Bottom - FullWind.Top
 .Right = FullWind.Right - FullWind.Left
 Rectangle hdc, .Left, .Top, .Right, .Bottom      ' Really easy to understand - draw a rectangle, hDC, and coordinates
End With
'
' The Washing Up bits
'
' This is a very important part, as it releases memory that has been taken up.
' If we don't do this, windows crashes due to a memory leak.
' You probably get a blue screen (altohugh I'm not sure)
'
' Get back the old region
SelectObject hdc, OldRegion
' Return the previous ROP
SetROP2 hdc, OldROP
' Return to the previous brush
SelectObject hdc, OldBrush
' Return the previous pen
SelectObject hdc, OldPen
' Delete the Brush I created
DeleteObject Brush
' Delete the Pen I created
DeleteObject Pen
' Delete the region I created
DeleteObject Region
' Release the hDC back to window's resource pool
ReleaseDC Myhwnd, hdc
End Sub
Private Sub UnDraw()
'
' Note, this sub is almost identical to the other one, except it doesn't go looking
' for the hWnd, it accesses the old one. Also, it doesn't clear the form.
' Otherwise, it just draws on top of the old one with an invert pen.
' 2 inverts = original
'
' If there hasn't been a border drawn, then get out of here.
If BorderDrawn = False Then Exit Sub
' Now set it
BorderDrawn = False
' If there isn't a current hWnd, then exit.
' That's why in the mouseup event we get out, because otherwise a border would be draw
' around the old window
If Myhwnd = 0 Then Exit Sub
Dim Cursor As POINTAPI    ' Cursor position
Dim RetVal As Long      ' Dummy returnvalue
Dim hdc As Long        ' hDC that we're going to be using
Dim Pen As Long        ' Handle to a GDI Pen object
Dim Brush As Long       ' Handle to a GDI Brush object
Dim OldPen As Long      ' Handle to previous Pen object (to restore it)
Dim OldBrush As Long     ' Handle to previous brush object (to restore it)
Dim OldROP As Long      ' Value of the previous ROP
Dim Region As Long      ' Handle to a GDI Region object that I create
Dim OldRegion As Long     ' Handle to previous Region object for the hDC
Dim FullWind As RECT     ' the bounding rectangle of the window in screen coords
Dim Draw As RECT       ' The drawing rectangle
'
' Getting all of the ingredients ready
'
' Get the full Rect of the window in screen co-ords
GetWindowRect Myhwnd, FullWind
' Create a region with width and height of the window
Region = CreateRectRgn(0, 0, FullWind.Right - FullWind.Left, FullWind.Bottom - FullWind.Top)
' Create an hDC for the hWnd
' Note: GetDC retrieves the CLIENT AREA hDC. We want the WHOLE WINDOW, including Non-Client
' stuff like title bar, menu, border, etc.
hdc = GetWindowDC(Myhwnd)
' Save the old region
RetVal = GetClipRgn(hdc, OldRegion)
' Retval = 0: no region   1: Region copied  -1: error
' Select the new region
RetVal = SelectObject(hdc, Region)
' Create a pen
Pen = CreatePen(DrawStyleConstants.vbSolid, 6, 0)  ' Draw Solid lines, width 6, and color black
' Select the pen
' A pen draws the lines
OldPen = SelectObject(hdc, Pen)
' Create a brush
' A brush is the filling for a shape
' I need to set it to a null brush so that it doesn't edit anything
Brush = GetStockObject(NULL_BRUSH)
' Select the brush
OldBrush = SelectObject(hdc, Brush)
' Select the ROP
OldROP = SetROP2(hdc, DrawModeConstants.vbInvert)  ' vbInvert means, whatever is draw,
     ' invert those pixels. This means that I can undraw it by doing the same.
'
' The Drawing Bits
'
' Put a box around the outside of the window, using the current hDC.
' These coords are in device co-ordinates, i.e., of the hDC.
With Draw
 .Left = 0
 .Top = 0
 .Bottom = FullWind.Bottom - FullWind.Top
 .Right = FullWind.Right - FullWind.Left
 Rectangle hdc, .Left, .Top, .Right, .Bottom      ' Really easy to understand - draw a rectangle, hDC, and coordinates
End With
'
' The Washing Up bits
'
' This is a very important part, as it releases memory that has been taken up.
' If we don't do this, windows crashes due to a memory leak.
' You probably get a blue screen (altohugh I'm not sure)
'
' Get back the old region
SelectObject hdc, OldRegion
' Return the previous ROP
SetROP2 hdc, OldROP
' Return to the previous brush
SelectObject hdc, OldBrush
' Return the previous pen
SelectObject hdc, OldPen
' Delete the Brush I created
DeleteObject Brush
' Delete the Pen I created
DeleteObject Pen
' Delete the region I created
DeleteObject Region
' Release the hDC back to window's resource pool
ReleaseDC Myhwnd, hdc
End Sub
```

