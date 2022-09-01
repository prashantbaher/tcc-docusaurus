---
title: VBA String Basics
tags:   [VBA]
permalink: /vba/string-basic/
---

The `String` data type represents a series of characters. This topic introduces the basic concepts of strings in Visual Basic.

## String Variables

An instance of a string can be assigned a value that represents a series of characters as shown in below example:

```vb showlinenumbers showLineNumbers
Dim MyString As String
MyString = "This is an example of the String data type"
```

A `String` variable can also accept any expression that evaluates to a string as shown in below example:

```vb showlinenumbers showLineNumbers
Dim OneString As String
Dim TwoString As String
OneString = "one, two, three, four, five"
TwoString = OneString.Substring(5, 3) ' Output -> "two".

OneString = "1"
TwoString = OneString & "1" ' Output -> "11".
```

Any [literal](https://binged.it/2T4EH0s) that is assigned to a `String` variable must be enclosed in quotation marks (""). 

This means that a quotation mark ("") within a string cannot be represented by a quotation mark. 

For example, the following code causes a compiler error:

```vb showlinenumbers showLineNumbers
Dim myString As String

' This line would cause an error.
myString = "He said, "Look at this example!""
```

This code causes an *error* because the compiler terminates the string after the second quotation mark, and the remainder of the string is interpreted as code. 

This means compiler think `He said, ` is a string and `Look at this example!` as a VB code.

But we want compiler to know that we want `He said, "Look at this example!"` as a string value.

To solve this problem, Visual Basic interprets two quotation marks in a string literal as one quotation mark in the string. 

The following example shows the correct way to include a quotation mark in a string:

```vb showlinenumbers showLineNumbers
' The value of myString is: He said, "Look at this example!"
myString = "He said, ""Look at this example!"" "
```

In the preceding example, the *two quotation marks* before and after the word `Look` become *one quotation mark* in the string. 

## The Immutability of Strings

A string is *immutable*, which means its value cannot be changed once it has been created. 

However, this does not prevent us from assigning more than one value to a string variable as shown in below example:

```vb showlinenumbers showLineNumbers
Dim myString As String = "This string is immutable"
myString = "Or is it?"
```

Here, a `string` variable is created, given a value, and then its value is changed.

In the first line, an instance of type `String` is created and given the value *This string is immutable*. 

In the second line of the example, a new instance is created and given the value *Or is it?*, and the string variable discards its reference to the first instance and stores a reference to the new instance.

Unlike other intrinsic data types, `String` is a reference type. 

When a variable of reference type is passed as an argument to a function or subroutine, a reference to the memory address where the data is stored is passed instead of the actual value of the string. 

So in the previous example, the name of the variable remains the same, but it points to a new and different instance of the String class, which holds the new value.

Next post will be about ***VBA Assignment Statements And Operators***.

