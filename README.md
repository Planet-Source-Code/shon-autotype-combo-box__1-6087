<div align="center">

## Autotype Combo Box


</div>

### Description

This code was taken from O'Neil. It searches a combo box as the user types. O'Neil's code was modified to use the SendMessage API to search the combo box, which made it much faster. This is very fast, even with thousands of records in the combo box. Thank you O'Neil for the idea, and the well commented code!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shon.md)
**Level**          |Intermediate
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shon-autotype-combo-box__1-6087/archive/master.zip)

### API Declarations

```
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Integer, _
  ByVal lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
```


### Source Code

```
Private Sub Combo1_Change()
  Dim i As Integer
  Dim l As Long
  Dim strNewText As String
  ' Check to see if a search is required.
  If Not IgnoreTextChange And Combo1.ListCount > 0 Then
    l = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal CStr(Combo1.Text))
    strNewText = Combo1.List(l)
    If Len(Combo1.Text) <> Len(strNewText) Then
      ' Partial match found
      ' Avoid recursively entering this event
      IgnoreTextChange = True
      i = Len(Combo1.Text)
      ' Attach the full text from the list to what has
      ' already been entered. This technique preserves
      ' the case entered by the user.
      Combo1.Text = Combo1.Text & Mid$(strNewText, i + 1)
      ' Select the text that is auto-entered
      Combo1.SelStart = i
      Combo1.SelLength = Len(Mid$(strNewText, i + 1))
    End If
  Else
    ' The IgnoreTwextChange Flag is only effective for one
    ' Changed event.
    IgnoreTextChange = False
  End If
End Sub
Private Sub Combo1_GotFocus()
  ' Select existing text on entry to the combo box
  Combo1.SelStart = 0
  Combo1.SelLength = Len(Combo1.Text)
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
  ' If a user presses the "Delete" key, then the selected text
  ' is removed.
  If KeyCode = vbKeyDelete And Combo1.SelText <> "" Then
    ' Make sure that the text is not automatically re-entered
    ' as soon as it is deleted
    IgnoreTextChange = True
    Combo1.SelText = ""
    KeyCode = 0
  End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
  ' If a user presses the "Backspace" key, then the selected text
  ' is removed. Autosearch is not re-performed, as that would only
  ' put it straight back again.
  If KeyAscii = 8 Then
    IgnoreTextChange = True
    If Len(Combo1.SelText) Then
      Combo1.SelText = ""
      KeyAscii = 0
    End If
  End If
  'if user presses enter, select the listindex
  If KeyAscii = 13 Then
    Combo1.ListIndex = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, ByVal CStr(Combo1.Text))
  End If
End Sub
```

