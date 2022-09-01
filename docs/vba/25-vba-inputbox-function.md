---
title: VBA InputBox Function
tags:   [VBA]
permalink: /vba/inputbox-function/
---

The VBA's `InputBox` function is useful for obtaining a single piece of information from the user. 

That information could be *a value*, *a text string*, or even *a range address*. 

This is a good alternative to developing a `UserForm` when you need to get only one value.

## InputBox syntax

Here’s a simplified version of the syntax for the *InputBox* function:

```vb showlinenumbers showLineNumbers
' InputBox syntax
InputBox(prompt[, title][, default])
```

The InputBox function accepts the arguments listed in below.

<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">InputBox Function Arguments</th>
    </tr>
    <tr>
        <th>Arguments</th>
        <th>What it means</tr>
    </tr>
    <tr>
        <td>prompt</td>
        <td>
            The text displayed in the input box.
        </td>
    </tr>
    <tr>
        <td>Title</td>
        <td>
            The text displayed in the input box’s title bar (optional).
        </td>
    </tr>
    <tr>
        <td>Default</td>
        <td>
            The default value for the user’s input (optional)
        </td>
    </tr>
</table>
-->

|Arguments|What it means|
|--- |--- |
|prompt|The text displayed in the input box.|
|Title|The text displayed in the input box’s title bar (optional).|
|Default|The default value for the user’s input (optional)|


## An InputBox example

Here’s an example showing how you can use the *InputBox* function:

```vb showlinenumbers showLineNumbers
' InputBox example
TheName = InputBox("What is your name?", "Greetings")
```

When you execute this VBA statement, application displays the dialog box shown in below figure. 

Notice that this example uses only the first two arguments and does not supply a default value. 

When the user enters a value and clicks `OK`, the routine assigns the value to the variable `TheName`.

![A-Simple-Message-Box](/assets/vba-images/Dialog_Boxes/InputBoxDialogBox.PNG)

Please note that VBA’s *InputBox function* always returns a `string`, so if you need to get a value, your code will need to do some additional checking. 

The following example uses the *InputBox function* to get a number. 

It uses the `IsNumeric` function to check whether the *string* is a *number*. 

If the string does contain a number, all is fine. 

If the user’s entry cannot be interpreted as a number, the code displays a message box.

```vb showlinenumbers showLineNumbers
' InputBox example
Sub GetDrawingSheetNumber()
  Dim NumberOfSheets as String
  Prompt = "How many sheets drawing have?"
  NumberOfSheets = InputBox (Prompt)

  If NumberOfSheets = "" Then Exit Sub
  If (IsNumeric)NumberOfSheets Then
    '......[Some code here]....
    Else
    MsgBox "Please enter a number."
  End If
End Sub
```

