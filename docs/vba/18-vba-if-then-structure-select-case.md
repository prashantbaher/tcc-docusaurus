---
title: If-Then-Else and Select Case structure
tags:   [VBA]
permalink: /vba/if-then-structure-select-case/
---

`If-Then` is VBA’s most important control structure. You’ll probably use this command on a daily basis. 

As in many other aspects of life, effective *decision-making* is the key to success in writing CAD or any other macros.

The `If-Then` structure has this basic syntax: 

```vb showLineNumbers
If condition Then statements [Else elsestatements]
```

Use the **If-Then** structure when you want to execute one or more statements conditionally. 

The optional **Else** clause, if included, lets you execute one or more statements if the condition you’re testing is not true. 

Sound confusing? Don’t worry; a few examples make this crystal clear. 

## If-Then examples

The following routine demonstrates the `If-Then` structure without the optional Else clause: 

```vb showLineNumbers
Sub GoodMorning()
  If Time < 0.5 Then MsgBox “Good Morning.”
End Sub
```

The `GoodMorning` procedure uses VBA’s `Time` function to get the system time. 

If the current system time is less than *.5* (in other words, before noon), the routine displays a friendly greeting. 

If Time is greater than or equal to *.5*, the routine ends and nothing happens. 

To display a different greeting if Time is greater than or equal to *.5*, add another *If-Then* statement after the first one: 

```vb showLineNumbers
Sub GoodMorning2()
  If Time < 0.5 Then MsgBox “Good Morning.”
  If Time >= 0.5 Then MsgBox “Good Afternoon.”
End Sub
```

Notice that I used `>=` (greater than or equal to) for the second *If-Then* statement. 

This ensures that the entire day is covered. Had I used `>` (greater than), then no message would appear if this procedure were executed at precisely 12:00 noon. 

## If-Then-Else examples

Another approach to the preceding problem uses the `Else` clause. 

Here’s the same routine recoded to use the `If-Then-Else` structure: 

```vb showLineNumbers
Sub GoodMorning3()
  If Time < 0.5 Then MsgBox “Good Morning.” Else _
  MsgBox “Good Afternoon.”
End Sub
```

Notice that I use the *line continuation character (underscore)* in the preceding example. 

The **If-Then-Else** statement is actually a single statement. 

VBA provides a slightly different way of coding **If-Then-Else** constructs that use an **End-If** statement. 

Therefore, the `GoodMorning` procedure can be rewritten as: 

```vb showLineNumbers
Sub GoodMorning4()
  If Time < 0.5 Then
    MsgBox “Good Morning.”
  Else
    MsgBox “Good Afternoon.”
  End If
End Sub
```

In fact, you can insert any number of statements under the `If` part, and any number of statements under the `Else` part. 

I prefer to use this syntax because it’s easier to read and makes the statements shorter. 

What if you need to expand the `GoodMorning` routine to handle three conditions: *morning, afternoon, and evening*?

You have two options: Use three `If-Then` statements or use a nested `If-Then-Else` structure. 

**Nesting** means placing an **If-Then-Else** structure within another **If-Then-Else** structure. 

The first approach, the three statements, is simplest: 

```vb showLineNumbers
Sub GoodMorning5()
  If Time < 0.5 Then Msg = “Morning.”
  If Time >= 0.5 And Time < 0.75 Then Msg = “Afternoon.”
  If Time >= 0.75 Then Msg = “Evening.”
  MsgBox “Good” & Msg
End Sub
```

The `Msg` variable gets a different text value, depending on the time of day. 

The final `MsgBox` statement displays the greeting: *Good Morning, Good Afternoon, or Good Evening*. 

The following routine performs the same action but uses an **If-Then-End If** structure: 

```vb showLineNumbers
Sub GoodMorning6()
  Dim Msg As String
  If Time < 0.5 Then
    Msg = “Morning.”
  If Time >= 0.5 And Time < 0.75 Then
    Msg = “Afternoon.”
  If Time >= 0.75 Then
    Msg = “Evening.”
  End If
  MsgBox “Good” & Msg
End Sub
```

## If-ElseIf-Else examples

In the previous examples, every statement in the routine is executed — even in the morning. 

A more efficient structure would exit the routine as soon as a condition is found to be true. 

In the morning, for example, the procedure should display the Good Morning message and then exit — without evaluating the other *superfluous* conditions. 

With a tiny routine like this, you don’t have to worry about execution speed. 

But for larger applications in which speed is important, you should know about another syntax for the If-Then structure. 

The `ElseIf` syntax follows: 

```vb showLineNumbers
If condition Then
[statements]
[Else condition-n Then
[elseifstatements]]
[Else
[elsestatements]]
```

Here’s how you can rewrite the `GreetMe` routine by using this syntax: 

```vb showLineNumbers
Sub GoodMorning7()
  Dim Msg As String
  If Time < 0.5 Then
    Msg = “Morning.”
  ElseIf Time >= 0.5 And Time < 0.75 Then
    Msg = “Afternoon.”
  Else
    Msg = “Evening.”
  End If
  MsgBox “Good” & Msg
End Sub
```

When a condition is `true`, VBA executes the conditional statements and the If structure ends. 

In other words, VBA doesn’t waste time evaluating the extraneous conditions, which makes this procedure a bit more efficient than the previous examples. 

The trade-off (there are always trade-offs) is that the code is more difficult to understand. (Of course, you already knew that.) 

# Select Case structure

The `Select` Case structure is useful for decisions involving three or more options (although it also works with two options, but using **If-Then-Else** structure is more efficient for that). 

The syntax for the `Select` Case structure follows: 

```vb showLineNumbers
Select Case testexpression
[Case expressionlist-n 
  [statements-n]]
[Case Else
  [elsestatements]]
End Select
```

Don’t be scared off by this official syntax. Using the **Select Case structure** is quite easy. 

## Select Case example

The following example shows how to use the **Select Case structure**. 

This also shows another way to code the examples presented in the previous section: 

```vb showLineNumbers
Sub SelectPartLength()
  Dim PartNumber As Integer
  Dim PartLength As Integer
  PartNumber = InputBox(“Please Enter part number:”)
  Select Case PartNumber
    Case Part001
      PartLength = 1
    Case Part002
      PartLength = 2
    Case Part003
      PartLength = 3
  End Select
  MsgBox “Part Length for this” & PartNumber & “is” & PartLength
End Sub
```

In this example, the `PartNumber` variable is being evaluated. 

The routine is checking for three different cases. 

Next post will be about ***VBA Looping***.
