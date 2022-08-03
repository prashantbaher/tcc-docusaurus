---
title: VBA Looping
tags:   [VBA]
permalink: /vba/looping/
---

The term *looping* refers to repeating a block of VBA statements numerous times. 

VBA provides various looping command for repeating code to make correct decision making. 

We will go through them in following topics: 

## For -Next Loop

The simplest type of loop is a `For-Next` loop. Here’s the syntax for this structure:

```vb
For counter = start To end [Step stepval]
[statements]
[Exit For]
[statements]
Next [counter]
```

The *looping* is controlled by a counter variable, which starts at one value and stops at another value. 

The statements between the `For` statement and the `Next` statement are the statements that get repeated in the loop. 

To see how this works, keep reading. 

### For-Next example

The following example shows a `For-Next` loop that doesn’t use the optional Step value or the optional `Exit` For statement. 

This routine loops 10 times and uses the VBA `MsgBox` function to show a number from 1 to 10: 

```vb
Sub ShowNumbers1()
  Dim i As Integer
  For i = 1 to 10
    MsgBox i
  Next i
End Sub
```

In this example, `i` (the loop counter variable) starts with a value of 1 and increases by 1 each time through the loop. 

Because I didn’t specify a Step value the `MsgBox` method uses the value of i as an argument. 

The first time through the *loop*, `i` is 1 and the procedure shows a number. 

The second time through (i = 2), the procedure show a number, and so on. 

```vb
Sub ShowNumbers2()
  Dim i As Integer Step 2
  For i = 1 to 10
    MsgBox i
  Next i
End Sub
```

Count starts out as 1 and then takes on a value of 3, 5, 7, and 9. The final Count value is 9. 

The Step value determines how the counter is *incremented*. Notice that the upper loop value (9) is not used because the highest value of Count after 9 would be 11, and 11 is larger than 10. 

### For-Next example with an Exit For statement 

A `For-Next` loop can also include one or more `Exit For` statements within the loop. 

When VBA encounters this statement, the loop terminates immediately. 

Here’s the same procedure as in the preceding section, rewritten to insert random numbers. 

```vb
Sub ShowNumbers3()
  Dim i As Integer Step 2
  For i = 1 to 10
    If i = 5 Then
      MsgBox “This is a mid value”
      Exit For
    End If
    MsgBox i
  Next i
End Sub
```

This routine performs the as earlier but when the variable i reached to 5, it shows a message, stating that this is a mid value and exit from loop. 

## Do-While Loop

VBA supports another type of looping structure known as a `Do-While` loop. 

Unlike a For-Next loop, a `Do-While` loop continues until a specified condition is met. 

Here’s the `Do-While` loop syntax:

```vb
' Do-While Structure
Do [While condition]
  [statements]
  [Exit Do]
  [statements]
Loop
```

The following example uses a `Do-While` loop. This routine uses 1 as a starting point and runs through next numbers. 

The loop continues until the routine encounter the condition of `i = 8`. 

```vb
' Do-While Example
Sub ShowNumbers4()
  Dim i As Integer
  Do While i <> 8
    MsgBox i
    i = i + 1
  Loop
End Sub
```

Some people prefer to code a `Do-While` loop as a `Do-Loop While` loop. 

This example performs exactly as the previous procedure but uses different loop syntax:

```vb
' Do-Loop While Example
Sub ShowNumbers5()
  Dim i As Integer
  Do 
    MsgBox i
    i = i + 1
  Loop While i <> 8
End Sub
```

Here’s the key difference between the `Do-While` and `Do-Loop While` loops. 

The `Do-While` loop always performs its conditional test first. If the test is not true, the instructions inside the loop are never executed. 

The `Do-Loop While` loop, on the other hand, always performs its conditional test after the instructions inside the loop are executed. 

Thus, the loop instructions are always executed at least once, regardless of the test. 

This difference can sometimes have a big effect on how your program functions. 

## Do-Until Loop

The `Do-Until` loop structure is similar to the `Do-While` structure. 

The two structures differ in their handling of the tested condition. 

A program continues to execute a `Do-While` loop while the condition remains true. 

In a `Do-Until` loop, the program executes the loop until the condition is true. Here’s the `Do-Until` syntax: 

```vb
' Do-Until Structure
Do [Until condition]
  [statements]
  [Exit Do]
  [statements]
Loop
```

The following example is the same one presented for the `Do-While` loop but recoded to use a `Do-Until` loop: 

```vb
Sub ShowNumbers6()
  Dim i As Integer
  Do Until i <> 8
    MsgBox i
    i = i + 1
  Loop
End Sub
```

Just like with the `Do-While` loop, you may encounter a different form of the `Do-Until` loop — a `Do-Loop Until` loop. 

The following example, which has the same effect as the preceding procedure, demonstrates an alternate syntax for this type of loop: 

```vb
Sub ShowNumbers7()
' Do-Loop Until Example
  Dim i As Integer
  Do 
    MsgBox i
    i = i + 1
  Loop Until i <> 8
End Sub
```

There is a subtle difference in how the `Do-Until` loop and the `Do-Loop Until` loop operate. 

In the former, the test is performed at the beginning of the loop, before anything in the body of the loop is executed. 

This means that it is possible that the code in the loop body will not be executed if the test condition is met. 

In the later version, the condition is tested at the end of the loop. 

Therefore, at a minimum, the `Do-Loop` Until loop always results in the body of the loop being executed once. 

Another way to think about it is like this: The `Do-While` loop keeps looping as long as the condition is true. 

The `Do-Until` loop keeps looping as long as the condition is False. 

## Looping through a Collection

VBA supports yet another type of looping — looping through each object in a **collection** of objects. 

Please note that I have not covered Object topic so far. For your understanding I give a brief explanation about **collection**. 

A **collection** is a group of same type of objects. 

For example, a drawing file in any CAD application is a collection of Sheets, and each sheet is a collection of drawing views and so on. 

When you need to loop through each object in a collection, use the For Each-Next structure. The syntax is 

```vb
' For Each-Next Structure
For Each element In collection
  [statements]
  [Exit For]
  [statements]
Next [element]
```

The following example loops through each drawing sheet in the active drawing and shows name of each active drawing sheet: 

```vb
' For Each-Next Example
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swPart As SldWorks.ModelDoc2
Dim swDwg As SldWorks.DrawingDoc
Dim BoolStatus As Boolean
Dim SheetNamesList As Variant
Sub ShowSheetName()
  Set swApp = Application.SldWorks
  Set swPart = swApp.ActiveDoc
  Set swDwg = swPart
  SheetNamesList = swDwg.GetSheetNames
  Dim SheetName As Variant
  For Each SheetName In SheetNamesList
    MsgBox SheetName
  Next SheetName
End Sub
```

In this example, first we get the list of all sheet names in opened drawing, then we loop through each sheet name in collection and show sheet name in a message box. 

For this example please notes that we did not need to load all sheet, this code can work on non-activate and non-loaded sheets also. 

Next post will be about ***Bug Finding***.

