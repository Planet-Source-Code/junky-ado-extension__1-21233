Attribute VB_Name = "modSelectAll"
' Highlights the textbox when selected
Public Sub SelectAll(ByRef txtBox As TextBox)
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
End Sub
