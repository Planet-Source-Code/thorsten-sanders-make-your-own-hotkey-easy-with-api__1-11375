<div align="center">

## Make your own Hotkey easy with API


</div>

### Description

The code is very easy to make your own systemwide hotkeys with API Call!
 
### More Info
 
Make a new module and put the API Call in the module.

You have to add a timer to your Form leave the standard name, put the intervall to 100 or when you want less or more you can do also, it is for how often he checks if the hotkey was pressed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thorsten Sanders](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thorsten-sanders.md)
**Level**          |Beginner
**User Rating**    |4.0 (28 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thorsten-sanders-make-your-own-hotkey-easy-with-api__1-11375/archive/master.zip)

### API Declarations

```
Public Declare Function GetAsyncKeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer
```


### Source Code

```
Private Sub Timer1_Timer()
If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyO) Then
MsgBox "It works :)"
End If
End Sub
'this example use the Control Key and O key as hotkey but you can use that key and how many keys you want alle the key codes you will find in the vb help under key code constants
```

