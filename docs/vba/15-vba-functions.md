---
title: VBA Functions
tags:   [VBA]
permalink: /vba/functions/
---

A `function` essentially performs a calculation and returns a single value. 

The `SUM` function in **MS Excel** returns the sum of a range of values. 

The same holds true for functions used in your **VBA expressions**: Each function does its thing and returns a single value.

The functions you use in VBA can come from two sources:

* Built-in functions provided by VBA
* Custom functions that you (or someone else) write, using VBA.

## Built-In VBA Functions

VBA provides numerous *built-in* functions. Some of these functions take arguments and some do not.

I present a few examples of VBA functions in code. 

In many of these examples, I use the `MsgBox` function to display a value in a message box. 

Yes, `MsgBox` is a VBA function — a rather unusual one, but a function nonetheless. 

This useful function displays a message in a pop-up dialog box. 

## Displaying the system date or time

The first example uses VBA’s `Date` function to display the current system date in a message box:

```vb
Sub ShowDate()
  MsgBox Date
End Sub
```

Notice that the `Date` function doesn’t use an argument. 

A VBA function with no argument doesn’t require an empty set of parentheses. 

In fact, if you type an empty set of parentheses, the VBE will promptly remove them.

To get the system time, use the `Time` function. And if you want it all, use the `Now` function to return both the date and the time. 

```vb
Sub ShowDate()
  MsgBox Now
End Sub
```

## Finding a string length

The following procedure uses the VBA's `Len` function, which returns the length of a text string. 

The `Len` function takes one argument: the `string`. 

When you execute this procedure, the *message box* displays **11** because the argument has **11** characters. 

```vb
Sub StringLength()
  Dim MyString As String
  Dim StringLength As Integer
  MyString = “Hello World”
  StringLength = Len(MyString)
  MsgBox StringLength
End Sub
```

## Displaying the integer part of a number

The following procedure uses the `Fix` function, which returns the integer portion of a value — *the value without any decimal digits*: 

```vb
Sub GetIntegerPart()
  Dim MyValue As Double
  Dim IntValue As Integer
  MyValue = 123.456
  IntValue = Fix(MyValue)
  MsgBox IntValue
End Sub
```

In this case, the message box displays **123**.

VBA has a similar function called `Int` Function. 

The difference between `Int` and `Fix` is how each deals with negative numbers. 

It’s a subtle difference, but sometimes it’s important. 

`Int` Function returns the first negative integer that’s less than or equal to the argument. `Int(-123.456)` returns **-124**. 

`Fix` Function returns the first negative integer that’s greater than or equal to the argument. `Fix(-123.456)` returns **-123**. 

## Determining a file size

The following `Sub` procedure displays the size, in bytes, of the executable file. 

It finds this value by using the `FileLen` function. 

```vb
Sub GetFileSize()
  Dim TheFile As String
  TheFile “C:\ProgramFiles\Program File\SolidworksCorp\SLDWORKS\SLDWORKS.exe”
  MsgBox FileLen(TheFile)
End Sub
```

Notice that this routine hard codes the filename (that is, it explicitly states the path). 

Generally, this **isn’t** a good idea. The file might not be on the *C drive*, or the Program File folder may have a different location. 

The following statement shows a better approach: 

```vb
TheFile = Application.Path & “\SLDWORKS.EXE” 
```

Path is a property of the Application object. 

It simply returns the name of the folder in which the application (that is, *Solidworks*) is installed (without a trailing backslash). 

## Identifying the type of a selected object

The following procedure uses the `TypeName` function, which returns the type of the selection (as a `string`): 

```vb
Sub ShowSelectionType()
  Dim SelType As String
  SelType = TypeName(Selection)
  MsgBox SelType
End Sub
```

This could be *a Sketch, a Part, a Assembly* or any *other type* of object that can be selected.

The `TypeName` function is very versatile. You can also use this function to determine the data type of a variable. 

Next post will be about ***VBA Functions that do more***.
