---
title: Programming Concepts, Comments and Data-types
tags:   [VBA]
permalink: /vba/programming-concepts/
---

`Visual Basic for application` (VBA) is a real, live programming language, it uses many elements common to all programming languages. 

In subsequent topics, we go through several of these elements: 

* Comments
* Variables
* Constants
* Data types
* Arrays and few other concepts

If you’ve programmed with other languages, some of this topics will be familiar or if you’re a programming *newbie*, it’s time to roll up your sleeves and get busy. 

## Comments in VBA Code

A `comment` is the simplest type of `VBA statement` because `VBA` ignores these statements, they can consist of anything you want. 

You can insert a `comment` to *remind* yourself why you did something or to *clarify* some particularly code you wrote. 

Use `comments` liberally and extensively to describe what the code does (which isn’t always obvious by reading the code itself). 

Often, code that makes perfect sense today mystifies you tomorrow.

You begin a `comment` with an apostrophe (*'*). `VBA` ignores any text that follows an apostrophe in a line of code.

You can use a complete line for your `comment` or insert your `comment` at the end of a line of code. 

The following example shows a `VBA` procedure with three `comments`, although they’re not necessarily good comments: 

```vb showlinenumbers showLineNumbers
Sub CommentDemo()
'  This procedure does nothing of value
   x = 0   'x represents nothing
   'Display the result
   MsgBox x
End Sub
```

The *apostrophe* indicates a `comment` rule has one exception. 

`VBA` doesn’t interpret an *apostrophe* inside a set of quotation marks as a `comment` indicator. 

For example, the following statement doesn’t contain a `comment`, even though it has an *apostrophe*: 

```vb showlinenumbers showLineNumbers
MsgBox = "Can't continue."
```

When you’re writing code, you may want to test a procedure by excluding a particular statement or group of statements. 

You could *delete* the statements and then *retype* them later but that’s a waste of time. 

A *better* solution is to simply turn those statements into `comments` by inserting *apostrophes*. 

`VBA` ignores statements beginning with *apostrophes* when executing a routine. 

To reactivate those `commented` statements, just remove the *apostrophes*.

Here’s a quick way to convert a block of statements to `comments`. 

In the `VBE`, choose **View -> Toolbars -> Edit** to display the *Edit* toolbar. 

To convert a block of statements to `comments`, select the statements and click the `Comment` Block button. 

To remove the *apostrophes*, select the statements and click the `Uncomment` Block button.

Although `comments` can be helpful, not all comments are created equal. 

For example, the following procedure uses lots of `comments`, but they add nothing of value. 

In this case, the code is clear enough without the comments. 

```vb showlinenumbers showLineNumbers
Sub BadComments()
'  Declare variables
   Dim x As Integer
   Dim y As Integer
   Dim z As Integer
'  Start the routine
   x = 100    'Assign 100 to x
   y = 100    'Assign 100 to y
'  Add x and y and store in z
   z = x + y
'  Show the result
   MsgBox z
End Sub
```

Everyone develops his or her own style of commenting. 

To be useful, however, `comments` should convey information that’s not immediately obvious from reading the code.

Otherwise, `comments` just chew up bytes and make files larger than necessary.

The following tips can help you make effective use of comments:

* Briefly describe the purpose of each `Sub` or `Function` procedure you write.

* Use *comments* to keep track of changes you make to a procedure.

* Use a *comment* to indicate that you’re using a function or a construct in an unusual or nonstandard manner.

* Use *comments* to describe the variables you use, especially if you don’t use meaningful variable names.

* Use a *comment* to describe any workarounds you develop to overcome bugs in your code.

* Write *comments* as you develop code, instead of saving the task for a final step.

* Depending on your work environment, consider adding a joke or two as a comment. The person who takes over your job when you get promoted might appreciate the humour.

## Data types in VBA

When I talk about `data type`, I’m referring to the manner in which a program stores data in memory - for example, as `integers`, `real numbers`, or `strings`. 

Although `VBA` can take care of these details automatically, it does so at a cost (There’s no free lunch). 

Letting `VBA` handle your data typing results in *slower execution* and *inefficient memory* use. 

For *small applications*, this usually doesn’t present much of a problem. 

But for *large or complex applications*, which may be slow or need to conserve every last byte of memory, you need to be on familiar terms with `data types`.

`VBA` automatically handles all the *data* details, which makes life easier for programmers. 

Not all programming languages provide this luxury. 

For example, some languages are strictly typed, which means the programmer must explicitly define the data type for every variable used.

`VBA` does not require that you declare the variables that you use, but it’s definitely a good practice. 

`VBA` has a variety of built-in *data types*. Below table lists the most common types of data that VBA can handle. 


<table class="w3-table-all w3-mobile  w3-card-4">
    <tr>
        <th class="w3-center" colspan="3">VBA’s Built-In Data Types</th>
    </tr>
    <tr>
        <th>Data Type</th>
        <th>Byte used</th>
        <th>Range of values</th>
    </tr>
    <tr>
        <td>Boolean</td>
        <td>2</td>
        <td>True or False</td>
    </tr>
    <tr>
        <td>Integer</td>
        <td>2</td>
        <td>–32,768 to 32,767</td>
    </tr>
    <tr>
        <td>Long</td>
        <td>4</td>
        <td>–2,147,483,648 to 2,147,483,647</td>
    </tr>
    <tr>
        <td>Single</td>
        <td>4</td>
        <td>–3.402823E38 to 1.401298E45</td>
    </tr>
    <tr>
        <td>Double (negative)</td>
        <td>8</td>
        <td>–1.79769313486232E308 to –4.94065645841247E-324</td>
    </tr>
    <tr>
        <td>Double (positive)</td>
        <td>8</td>
        <td>4.94065645841247E–324 to  1.79769313486232E308</td>
    </tr>
    <tr>
        <td>Currency</td>
        <td>8</td>
        <td>–922,337,203,685,477.5808 to 922,337,203,685,477.5807</td>
    </tr>
    <tr>
        <td>Date</td>
        <td>8</td>
        <td>1/1/100 to 12/31/9999</td>
    </tr>
    <tr>
        <td>String</td>
        <td>1 per char</td>
        <td>varies</td>
    </tr>
    <tr>
        <td>Object</td>
        <td>4</td>
        <td>Any defined object</td>
    </tr>
    <tr>
        <td>Variant</td>
        <td>Varies</td>
        <td>Any data type</td>
    </tr>
    <tr>
        <td>User defined</td>
        <td>Varies</td>
        <td>Varies</td>                    
    </tr>
</table>

Next post will be about ***VBA Variables***.
