<div align="center">

## AutoResize


</div>

### Description

This code resizes a form and its controls according to the screen resolution. It also takes into account the size of the screen fonts (although this is untested!).
 
### More Info
 
designwidth - the width that your app was designed at (i.e. 800 or 1024)

designheight - the height that your app was designed at (i.e. 600 or 768)

designfontsize - the size of the screen fonts (small - 96, large - 120)

The function to resize depending upon the size of the fonts is untested as yet because my PC keeps crashing if I change the font size. If I doesn't work then could you let me know.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Parter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-parter.md)
**Level**          |Unknown
**User Rating**    |3.7 (22 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-parter-autoresize__1-2369/archive/master.zip)

### API Declarations

```
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
```


### Source Code

```
'Place the following line in the Form_Load procedure of the form
AutoResize Me, 2 'put a 2 for a full screen form or a 0 for any other form
'Place the following in a module
Sub AutoResize(frmName As Form, winstate As Integer)
Dim designwidth As Integer, designheight As Integer, designfontsize As Integer, currentfontsize As Integer
Dim ratiox As Single, ratioy As Single, numofcontrols As Integer, a As Integer
Dim fontratio As Single
'Change the designwidth and the designheight according to the resolution that the form was designed at
designwidth = 1024
designheight = 768
designfontsize = 96
'Get the current resolution
resx = Screen.Width / Screen.TwipsPerPixelX
resy = Screen.Height / Screen.TwipsPerPixelY
'Work out the ratio for resizing the controls
ratiox = resx / designwidth
ratioy = resy / designheight
'check to see what size of fonts are being used
If IsScreenFontSmall Then
  currentfontsize = 96
Else
  currentfontsize = 120
End If
'work out the ratio for the fontsize
fontratio = currentfontsize / designfontsize
If ratiox = 1 And ratioy = 1 And fontratio = 1 Then Exit Sub
numofcontrols = frmName.Controls.Count - 1
For a = 0 To numofcontrols
  If TypeOf frmName.Controls(a) Is CommandButton Then
    frmName.Controls(a).Width = frmName.Controls(a).Width * ratiox
    frmName.Controls(a).Height = frmName.Controls(a).Height * ratioy
    frmName.Controls(a).Top = frmName.Controls(a).Top * ratioy
    frmName.Controls(a).Left = frmName.Controls(a).Left * ratiox
    frmName.Controls(a).FontSize = frmName.Controls(a).FontSize * ratiox
  ElseIf TypeOf frmName.Controls(a) Is Timer Then
  Else
    frmName.Controls(a).Width = frmName.Controls(a).Width * ratiox
    frmName.Controls(a).Height = frmName.Controls(a).Height * ratioy
    frmName.Controls(a).Top = frmName.Controls(a).Top * ratioy
    frmName.Controls(a).Left = frmName.Controls(a).Left * ratiox
  End If
Next a
If fontratio <> 1 Then
  For a = 0 To numofcontrols
    If TypeOf frmName.Controls(a) Is CommandButton Then
      frmName.Controls(a).Width = frmName.Controls(a).Width * fontratio
      frmName.Controls(a).Height = frmName.Controls(a).Height * fontratio
      frmName.Controls(a).Top = frmName.Controls(a).Top * fontratio
      frmName.Controls(a).Left = frmName.Controls(a).Left * fontratio
      frmName.Controls(a).FontSize = frmName.Controls(a).FontSize * fontratio
    ElseIf TypeOf frmName.Controls(a) Is Timer Then
    Else
      frmName.Controls(a).Width = frmName.Controls(a).Width * fontratio
      frmName.Controls(a).Height = frmName.Controls(a).Height * fontratio
      frmName.Controls(a).Top = frmName.Controls(a).Top * fontratio
      frmName.Controls(a).Left = frmName.Controls(a).Left * fontratio
    End If
    Next a
End If
If winstate = 0 Then
  frmName.Height = frmName.Height * ratioy
  frmName.Width = frmName.Width * ratiox
ElseIf winstate = 2 Then
  frmName.Width = Screen.Width
  frmName.Height = Screen.Height
  frmName.Top = 0
  frmName.Left = 0
End If
End Sub
Public Function IsScreenFontSmall() As Boolean
Dim hWndDesk As Long
Dim hDCDesk As Long
Dim logPix As Long
Dim r As Long
hWndDesk = GetDesktopWindow()
hDCDesk = GetDC(hWndDesk)
logPix = GetDeviceCaps(hDCDesk, LOGPIXELSX)
r = ReleaseDC(hWndDesk, hDCDesk)
If logPix = 96 Then IsScreenFontSmall = True
End Function
```

