Attribute VB_Name = "SubMain"
Sub main()
If Left$(Command$, 2) = "/c" Then Form2.Show: Exit Sub
If Left$(Command$, 2) = "/p" Then MsgBox "Text Rotater Screen Saver, By Marco Samy", vbExclamation: Exit Sub
Form1.Show
End Sub
