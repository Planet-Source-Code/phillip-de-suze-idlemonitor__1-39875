<div align="center">

## IdleMonitor


</div>

### Description

This detects your idlestate if there is no input from your keyboard or mouse a triggers an event

so if you want to make your screensavers this is jus what you need..
 
### More Info
 
Interval


<span>             |<span>
---                |---
**Submitted On**   |2002-09-19 22:35:02
**By**             |[Phillip De Suze](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/phillip-de-suze.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[IdleMonito14689210162002\.zip](https://github.com/Planet-Source-Code/phillip-de-suze-idlemonitor__1-39875/archive/master.zip)

### API Declarations

```
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
```





