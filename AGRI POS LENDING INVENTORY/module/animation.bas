Attribute VB_Name = "animation"
Option Explicit

Public Function coolClose(FormClose As Form, speed As Integer)
Do Until FormClose.Height <= 405
DoEvents
FormClose.Height = FormClose.Height - speed * 9
FormClose.Top = FormClose.Top + speed * 5
Loop
Do Until FormClose.Width <= 1680
DoEvents
FormClose.Width = FormClose.Width - speed * 9
FormClose.Left = FormClose.Left + speed * 5
Loop
Unload FormClose

End Function
