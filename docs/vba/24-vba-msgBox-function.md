---
title: VBA MsgBox Function
tags:   [VBA]
permalink: /vba/msgBox-function/
---

You’re probably already familiar with the VBA `MsgBox` function — I use it quite a bit in the examples. 

The `MsgBox` function, which accepts the arguments shown in below table, is handy for displaying information and getting simple user input. 

It’s able to get user input because it’s a function. 

A *function*, as you recall, returns a value. 

In the case of the `Msgbox` function, it uses a dialog box to get the value that it returns. 

Keep reading to see exactly how it works.

Here’s a simplified version of the syntax for the MsgBox function:

```vb showlinenumbers showLineNumbers
' MsgBox Structure
MsgBox(prompt[, buttons][, title])
```
<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">MsgBox Function Arguments</th>
    </tr>
    <tr>
        <th>Arguments</th>
        <th>What it does</tr>
    </tr>
    <tr>
        <td>prompt</td>
        <td>
            The text your application displays in the message box
        </td>
    </tr>
    <tr>
        <td>buttons</td>
        <td>
            A number that specifies which buttons (along with what icon) appear in the message box (optional)
        </td>
    </tr>
    <tr>
        <td>title</td>
        <td>
            The text that appears in the message box’s title bar (optional) displaying a simple message box
        </td>
    </tr>
</table>
-->

|Arguments|What it does|
|--- |--- |
|prompt|The text your application displays in the message box|
|buttons|A number that specifies which buttons (along with what icon) appear in the message box (optional)|
|title|The text that appears in the message box’s title bar (optional) displaying a simple message box|

You can use the *MsgBox* function in two ways:

* To simply show a message to the user. In this case, you don’t care about the result returned by the function.
* To get a response from the user. In this case, you do care about the result returned by the function. The result depends on the button that the user clicks.

If you use the *MsgBox* function by itself, don’t include parentheses around the arguments. 

The following example simply displays a message and does not return a result. 

When the message is displayed, the code stops until the user clicks `OK`.

```vb showlinenumbers showLineNumbers
' MsgBox function Example
Sub main()
  MsgBox "Hello, world!"
End Sub
```

Below figure shows how this message box looks:

![A-Simple-Message-Box](/assets/vba-images/Dialog_Boxes/ASimpleMessageBox.PNG)

## Getting a response from a message box

If you display a message box that has more than just an **OK** button, you’ll probably want to know which button the user clicks. 

The *MsgBox function* can return a value that represents which button is clicked. 

You can assign the result of the MsgBox function to a variable.

In the following code, I use some built-in constants that make it easy to work with the values returned by `MsgBox`:

```vb showlinenumbers showLineNumbers
' MsgBox built-in constants Example
Sub GetAnswer()
  Dim Ans as Integer
  Ans = MsgBox ("Did you eat lunch?", vbYesNo)
  Select Case Ans
    Case vbYes
    '......[Some code here]....
    Case vbNo
    '......[Some code here]....
  End Select
End Sub
```

Below figure shows how it looks. 

When you execute this procedure, the `Ans` variable is assigned a value of either `vbYes` or `vbNo`, depending on which button the user clicks. 

The `Select` Case statement uses the `Ans` value to determine which action the code should perform.

![A-Simple-Message-Box-with-two-buttons](/assets/vba-images/Dialog_Boxes/ASimpleMessageBoxWithTwoButtons.PNG)

You can also use the `MsgBox` function result without using a variable, as the following example demonstrates:

```vb showlinenumbers showLineNumbers
' MsgBox without variable
Sub GetAnswer2()
  If MsgBox ("Continue?", vbYesNo) = vbYes Then
  '......[Some code here]....
  Else
  '......[Some code here]....
  End If
End Sub
```

## Customizing message boxes

The flexibility of the buttons argument makes it easy to customize your message boxes. 

You can specify which buttons to display, determine whether an icon appears, and decide which button is the default (the default button is “clicked” if the user presses `Enter`).

Below table lists some of the built-in constants you can use for the buttons argument. 

If you prefer, you can use the value rather than a constant (but I think using the built-in constants is a lot easier).

<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="3">Constants Used in the MsgBox Function</th>
    </tr>
    <tr>
        <th>Constant</th>
        <th>Value</th>
        <th>What it does</th>
    </tr>
    <tr>
        <td>vbOKOnly</td>
        <td>0</td>
        <td>Display OK button only.</td>
    </tr>
    <tr>
        <td>vbOKCancel</td>
        <td>1</td>
        <td>Display OK and Cancel buttons</td>
    </tr>
    <tr>
        <td>vbAbortRetryIgnore</td>
        <td>2</td>
        <td>Displays Abort, Retry, and Ignore buttons.</td>
    </tr>
    <tr>
        <td>vbYesNoCancel</td>
        <td>3</td>
        <td>Displays Yes, No, and Cancel buttons.</td>
    </tr>
    <tr>
        <td>vbYesNo</td>
        <td>4</td>
        <td>Displays Yes and No buttons.</td>
    </tr>
    <tr>
        <td>vbRetryCancel</td>
        <td>5</td>
        <td>Displays Retry and Cancel buttons.</td>
    </tr>
    <tr>
        <td>vbCritical</td>
        <td>16</td>
        <td>Displays Critical Message icon.</td>
    </tr>
    <tr>
        <td>vbQuestion</td>
        <td>32</td>
        <td>Displays Warning Query icon.</td>
    </tr>
    <tr>
        <td>vbExclamation</td>
        <td>48</td>
        <td>Displays Warning Message icon.</td>
    </tr>
    <tr>
        <td>vbInformation</td>
        <td>64</td>
        <td>Displays Information Message icon.</td>
    </tr>
    <tr>
        <td>vbDefaultButton1</td>
        <td>0</td>
        <td>First button is default.</td>
    </tr>
    <tr>
        <td>vbDefaultButton2</td>
        <td>256</td>
        <td>Second button is default.</td>
    </tr>
    <tr>
        <td>vbDefaultButton3</td>
        <td>512</td>
        <td>Third button is default.</td>
    </tr>
    <tr>
        <td>vbDefaultButton4</td>
        <td>768</td>
        <td>Fourth button is default.</td>
    </tr>
</table>
-->

|Constant|Value|What it does|
|--- |--- |--- |
|vbOKOnly|0|Display OK button only.|
|vbOKCancel|1|Display OK and Cancel buttons|
|vbAbortRetryIgnore|2|Displays Abort, Retry, and Ignore buttons.|
|vbYesNoCancel|3|Displays Yes, No, and Cancel buttons.|
|vbYesNo|4|Displays Yes and No buttons.|
|vbRetryCancel|5|Displays Retry and Cancel buttons.|
|vbCritical|16|Displays Critical Message icon.|
|vbQuestion|32|Displays Warning Query icon.|
|vbExclamation|48|Displays Warning Message icon.|
|vbInformation|64|Displays Information Message icon.|
|vbDefaultButton1|0|First button is default.|
|vbDefaultButton2|256|Second button is default.|
|vbDefaultButton3|512|Third button is default.|
|vbDefaultButton4|768|Fourth button is default.|


For using more than one of these constants as an argument, just connect them with a `+` operator. 

For example, to display a message box with `Yes` and `No` buttons and an exclamation icon, use the following expression as the second *MsgBox* argument:

```vb showlinenumbers showLineNumbers
' Using multiple MsgBox built-in constants
vbYesNo + vbExclamation
```

Or, if you prefer to make your code less understandable, use a value of *52 (that is, 4 + 48)*.

The following example uses a combination of constants to display a message box with a `Yes button` and a `No button` (`vbYesNo`) as well as a question mark icon (`vbQuestion`). 

The constant `vbDefaultButton2` designates the second button (`No`) as the default button — that is, the button that is clicked if the user presses `Enter`. 

For simplicity, we assign these constants to the `Config` variable and then use `Config` as the second argument in the *MsgBox* function:

```vb showlinenumbers showLineNumbers
' Using multiple MsgBox built-in constants
Sub GetAnswer3()
  Dim Config As Integer
  Dim Ans as Integer
  Config = vbYesNo + vbQuestion + vbDefaultButton2
  Ans = MsgBox("Is part opened?", Config)
  If Ans = vbYes Then OpenPart
End Sub
```

Below figure shows the message box application displays when you execute the `GetAnswer3` procedure. 

If the user clicks the *Yes button*, the routine executes the procedure named `OpenPart` (which is not shown). 

If the user clicks the *No button* (or presses `Enter`), the routine ends with no action. 

Because I omitted the title argument in the *MsgBox* function, our application uses the default title, in my case it is *Solidworks*.

![MsgBox-function-button](/assets/vba-images/Dialog_Boxes/MsgBoxfunctionsbutton.PNG)

Previous examples have used constants (such as `vbYes` and `vbNo`) for the return value of a *MsgBox* function. 

Besides these two constants, below table lists a few others.

<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="3">Constants Used as Return Values for the MsgBox Function</th>
    </tr>
    <tr>
        <th>Constant</th>
        <th>Value</th>
        <th>What it does</th>
    </tr>
    <tr>
        <td>vbOK</td>
        <td>1</td>
        <td>User clicked OK.</td>
    </tr>
    <tr>
        <td>vbCancel</td>
        <td>2</td>
        <td>User clicked Cancel.</td>
    </tr>
    <tr>
        <td>vbAbort</td>
        <td>3</td>
        <td>User clicked Abort.</td>
    </tr>
    <tr>
        <td>vbRetry</td>
        <td>4</td>
        <td>User clicked Retry.</td>
    </tr>
    <tr>
        <td>vbIgnore</td>
        <td>5</td>
        <td>User clicked Ignore.</td>
    </tr>
    <tr>
        <td>vbYes</td>
        <td>6</td>
        <td>User clicked Yes.</td>
    </tr>
    <tr>
        <td>vbNo</td>
        <td>7</td>
        <td>User clicked No.</td>
    </tr>
</table>
-->

|Constant|Value|What it does|
|--- |--- |--- |
|vbOK|1|User clicked OK.|
|vbCancel|2|User clicked Cancel.|
|vbAbort|3|User clicked Abort.|
|vbRetry|4|User clicked Retry.|
|vbIgnore|5|User clicked Ignore.|
|vbYes|6|User clicked Yes.|
|vbNo|7|User clicked No.|

Next post will be about ***VBA InputBox Function***.

