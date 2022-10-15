'MIT License

'Copyright (c) 2021 Ricardo Camisa

'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

Public FontSize As Integer
Public FontName As String
Public fColor, eColor, tColor, bColor
Public cBox     As clsTextBox
Public ctBox    As New Collection
public ctrl     as Control
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


