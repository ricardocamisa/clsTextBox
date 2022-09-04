# VBA TextBox Styling

It is very simple to implement this class in your project

<a href="https://tukeya.herokuapp.com" target="_blank"><img src="https://miro.medium.com/max/700/0*KvBjxdkU_J5IXgXI" alt="TextBox" style="height: 250px !important;width: 450px !important;" ></a>

## Getting Started
> Importing or copying both **clsTextBox.cls** and **mdTextBox.bas** is **required** in order to work!

Here is a basic template, simply add this to a userform.

```vb
Public cBox as clsTextBox

Private Sub UserForm_Initialize()
    Set cBox = New clsTextBox
    cBox.AddEventListenerAll Me
End Sub
```




