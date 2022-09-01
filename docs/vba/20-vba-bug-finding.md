---
title: Bug Finding
tags:   [VBA]
permalink: /vba/bug-finding/
---

A **bug** is an error in your programming. Here I cover the topic of programming bugs — how to identify them and how to remove them from your module. 

## Types of Bugs

The term *program bug*, as you probably know, refers to a problem with software. 

In other words, if software doesn’t perform as expected, it has a **bug**. 

Fact is, all major software programs have bugs — lots of bugs. 

A CAD software like *Solidworks* itself has hundreds (if not thousands) of bugs. 

Fortunately, the vast majority of these bugs are relatively obscure and appear in only very specific circumstances.

When you write non-trivial VBA programs, your code probably will have bugs. 

This is a fact of life and not necessarily a reflection of your programming ability. The bugs may fall into any of the following categories: 

  * *Logical flaws in your code*: You can often avoid these bugs by carefully thinking through the problem your program addresses.

  * *Incorrect context bugs*: This type of bug surfaces when you attempt to do something at the wrong time. For example, you may try to update the sketch dimension and there are no sketch is activated. 

  * *Extreme-case bugs*: These bugs rear their heads when you encounter data you didn’t anticipate, such as very large or very small numbers. 

  * *Wrong data types bugs*: This type of bug occurs when you try to process data of the wrong type, such as attempting to take the square root of a text string. 

**Debugging** is the process of identifying and correcting bugs in your program. 

Developing debugging skills takes time, so don’t be discouraged if this process is difficult at first. 

It’s important to understand the distinction between bugs and syntax errors. 

A **syntax error** is a language error. For example, you might misspell a keyword, omit the Next statement in a `For-Next` loop, or have a mismatched parenthesis. 

Before you can even execute the procedure, you must correct these syntax errors. 

A program bug is much *subtler*. You can execute the routine, but it doesn’t perform as expected. 

## Identifying Bugs

Before you can do any debugging, you must determine whether a bug actually exists. 

You can tell that your macro contains a bug if it doesn’t work the way it should. Usually, but not always, you can easily discern this. 

A key fact known to all programmers is that bugs often appear when you least expect them. 

For example, just because your macro works fine with one data set doesn’t mean you can assume it will work equally as well with all data sets. 

Or your macro runs fine in your system but not working properly in your friend's system. 

Such cases happened all the time and are part of debugging. 

The best debugging approach is to start with thorough testing, under a variety of real-life conditions. 

Because any changes made by your VBA code cannot be undone, it is always a good idea to use a backup copy of your CAD files that you use for testing. 

I usually copy some files into a temporary folder and use those files for my testing. 

## Debugging Techniques

In this section, I discuss the some of the most common methods for debugging VBA code: 

* Examine your code
* Inserting `MsgBox` functions at various locations in your code
* Inserting `Debug.Print` statement

### Examine your code

Perhaps the most straightforward debugging technique is simply taking a close look at your code to see whether you can find the problem. 

If you’re lucky, the error jumps right out, and you can fix the problem.

Notice I said, “If you’re lucky.” That’s because often you discover errors when you have been working on your program for long hours and you are running on caffeine and willpower. 

At times like that, you are lucky if you can even see your code. 

Thus, don’t be surprised if simply examining your code isn’t enough to make you find and expunge all the bugs it contains. 

### Using the MsgBox function

A common problem in many programs involves one or more variables not taking on the values you expect. 

In such cases, monitoring the variable(s) while your code runs is a helpful debugging technique. 

One way to do this is by inserting temporary `MsgBox` functions into your routine. 

For example, I used `MsgBox` function to check conditions. Whenever I use `If-Else` statement, I put one message in `If` condition and put another message in `Else` condition. 

By this way, I make sure condition which I want is working correct or not. 

Feel free to use `MsgBox` functions frequently when you debug your code. 

Just make sure that you remove them after you identify and correct the problem.

### Inserting Debug.Print Statement

As an alternative to using `MsgBox` functions in your code, you can insert one or more temporary `Debug.Print` statements. 

Use these statements to print the value of one or more variables in the *Immediate* window. 

Here’s an example that displays a message of "This condition is working fine". 

```vb showlinenumbers showLineNumbers
If swPart Is Nothing Then
  Debug.Print "This condition is working fine."
```

If VBE’s **Immediate** window is not visible, press `Ctrl+G`.

Unlike `MsgBox`, `Debug.Print` statements do not halt your code. 

So you’ll need to keep an eye on the **Immediate** window to see what’s going on. 

After you’ve debugged your code, be sure to remove all the `Debug.Print` statements.

Next post will be about ***VBA Debugger***.

