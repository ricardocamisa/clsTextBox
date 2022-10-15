# Modern textbox and combobox UI/UX in VBA Userforms

clsTextBox is a VBA Project that allows you to leave the vba userfom textbox and comboboxes with a sleek appearance and reduces a lot of development time, while ensuring an excellent user experience.

Multiple forms are tracked simultaneously. Just call the `SetClassTextBox` for each form.

## Installation

Just import the following 2 code modules in your VBA Project:

- [**mdlTextBox.bas**](https://github.com/ricardocamisa/clsTextBox/blob/main/src/mdlTextBox.bas)
- [**clsTextBox.cls**](https://github.com/ricardocamisa/clsTextBox/blob/main/src/clsTextBox.cls)

## Usage

In your Modal Userform use:

```vba
SetClassTextBox Me
```

For example you can use your Form's Initialize Event:

```vba
Private Sub UserForm_Initialize()
    SetClassTextBox Me
End Sub
```

Opcionally you can add the colors to the background, text color before clicking, text color when entering the text, font size end font family, following the next example:

```VBA
Private Sub UserForm_Initialize()
    SetClassTextBox Me, 1512210, 14854934, 10395294, 16447476, 10, "MontSerrat"
End Sub
```

## Notes

- First color: TextBox ForeColor
- Second color: When TextBox Enter
- Third color: Title and bottom line Color
- Fourth color: Background Color
- Fifth: Font Size
- Sixth: Font Family

You can also use the RGB format to customize colors.

## ADD MouseIcon on your Forms

You can also include ico by hovering over the buttons with one of the `Tag`, `Name`, or `ControlTipText` properties as follows: `btn` for example: `btnSave`

## License

MIT License

Copyright (c) 2021 Ricardo Camisa

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
