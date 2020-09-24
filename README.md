<div align="center">

## A Very Flat Std Command Button


</div>

### Description

Make a standard command button very flat ;-)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ulli-a-very-flat-std-command-button__1-56117/archive/master.zip)





### Source Code

```
Option Explicit
Private Type RECT
  Left  As Long
  Top   As Long
  Right  As Long
  Bottom As Long
End Type
Private WindowRect  As RECT
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Sub Form_Load()
 Const SnippOff  As Long = 3
 Dim hRgn     As Long
  With WindowRect
    .Left = SnippOff
    .Top = SnippOff
    .Right = ScaleX(Command1.Width, ScaleMode, vbPixels) - SnippOff
    .Bottom = ScaleY(Command1.Height, ScaleMode, vbPixels) - SnippOff
    hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
    SetWindowRgn Command1.hWnd, hRgn, True
    DeleteObject hRgn
  End With
End Sub
```

