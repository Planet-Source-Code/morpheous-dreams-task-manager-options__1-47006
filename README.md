<div align="center">

## Task Manager Options


</div>

### Description

This has 5 options don't seem like alot but it could help someone new maybe? This can disable Task Manager / Enable it | Hide / Show it and last of all can close it... Simple but people hideing there program could use a simple call to close windows task manager instead of useing App.TaskVisible = False that leaves the program in the processes tree view of windows task manager. Maybe it will help someone/Maybe Not
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Morpheous Dreams](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/morpheous-dreams.md)
**Level**          |Beginner
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/morpheous-dreams-task-manager-options__1-47006/archive/master.zip)

### API Declarations

```
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const WM_CLOSE = &H10
```


### Source Code

```
'Enable windows task manager
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call EnableWindow(x, 1)
'...................................
'Disable windows task manager
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call EnableWindow(x, 0)
'....................................
'Hide windows task manager
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call ShowWindow(x, SW_HIDE)
'......................................
'Show windows task manager
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call ShowWindow(x, SW_SHOW)
'...................................
'Close windows task magager
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
'..........................................
```

