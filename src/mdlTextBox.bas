Attribute VB_Name = "mdlTextBox"
Public FontSize As Integer
Public FontName As String
Public fColor, eColor, tColor, bColor
Public cBox     As clsTextBox
Public ctBox    As New Collection
Public Sub SetClassTextBox(Form As MSForms.UserForm, _
                        Optional fColorValue As String = 1512210, _
                        Optional eColorValue As String = 14854934, _
                        Optional tColorValue As String = 10395294, _
                        Optional bColorValue As String = 16447476, _
                        Optional fFontSize As Double = 10, _
                        Optional fFontName As String = "MontSerrat")
                        fColor = fColorValue: eColor = eColorValue: tColor = tColorValue: bColor = bColorValue: FontSize = fFontSize: FontName = fFontName
                        Set cBox = New clsTextBox
                        cBox.Init Form: End Sub


