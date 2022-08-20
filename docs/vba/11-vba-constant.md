---
title: VBA Constants
tags:   [VBA]
permalink: /vba/constant/
---

A variable’s value may (and usually does) change while your procedure is executing. 

That’s why they call it a **variable**. 

Sometimes you need to refer to a **value** or **string** that never changes. 

In such a case, you need a **constant** — a named element whose value doesn’t change.

As shown in the following examples, you declare **constants** by using the `Const` statement. 

The declaration statement also gives the constant its value: 

```vb showLineNumbers
Const BlockLength As Integer = 4.
Const BlockThickness = .5
Const PartName As String = "Part Name:"
Public Const AppName As String = "Part Calculation"
```

Using *constants* in place of hard-coded *values* or *strings* is an excellent programming practice. 

For example, if your procedure needs to refer to a specific value (such as *sheet thickness*) several times. 

It is better to declare the value as a *constant* and refer to its *name* rather than the *value*. 

This makes your code more readable and easier to change. 

When sheet thickness changes, you have to change only one statement rather than several.

Like variables, constants have a scope. Keep these points in mind:

* To make a *constant* available within only a *single procedure*, declare the constant after the procedure’s `Sub` or `Function` statement. 
* To make a *constant* available to *all procedures* in a module, declare the constant in the **Declarations** section for the module. 
* To make a *constant* available to *all modules*, use the `Public` keyword and declare the constant in the **Declarations** section of any module. 

If you attempt to change the value of a constant in a **VBA routine**, you get an error. 

This isn’t too surprising because a Constant is `constant`. 

Unlike a variable, the value of a constant *does not* vary. 

If you need to change the value of a constant while your code is running, what you really need is a variable.

## Pre-made constants

Your **CAD Application** and **VBA** contain many predefined constants, which you can use without the need to declare them yourself. 

The macro recorder (in *Solidworks*) usually uses constants rather than actual values. 

In general, you don’t need to know the value of these constants to use them. 

The following simple procedure uses a **built-in** constant `swDefaultTemplatePart` to select the default part template while opening a new file. 

```vb showLineNumbers
set swPart = swApp.NewDocument(swApp.GetUserPreferenceStringValue _
    (swUserPreferenceStringValue_e.swDefaultTemplatePart),0,0,0)
```

In above example, *Solidworks* did not record these constants. 

It simply generates the direct path to open part document.

To find the actual value of a built-in constant, use the "Immediate window" in the VBE, and execute a VBA statement such as the following: 

```vb showLineNumbers
?swDefaultTemplatePart
```

> If the Immediate window isn’t visible, press `Ctrl+G`. The question mark is a shortcut for typing `Print`. 

Next post will be about ***VBA Strings Basics***.

