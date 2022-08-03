---
title: Sub and Function Procedures
tags:   [VBA]
permalink: /vba/sub-and-function-procedure/
---

The `VBA` code that you write in the `Visual Basic Editor` is known as a `procedure`. 

The two most common types of procedures are `Sub` and `Function`.

* A `Sub` procedure is a group of `VBA statements` that performs an *action* (or *actions*).
* A `Function` procedure is a group of `VBA statements` that performs a *calculation* and returns a *single value*.

Most of the macros you write in `VBA` are `Sub` procedures. 

You can think of a `Sub` procedure as being like a command: "Execute the `Sub` procedure and something happens". 

Exactly what happens depends on the `Sub` procedure’s `VBA code`.

A `Function` is also a procedure, but it’s quite different from a `Sub`. 

For understanding a `Function` let’s take an example from `MS Excel`. 

**Excel** includes many worksheet functions that you use every day. Examples include `SUM`. This `SUM function` takes input in form of values and then it does calculation behind the scene, and then returns a single value. 

The same goes for `Function procedure` that you develop with `VBA`.

## Sub procedures

Every `Sub` procedure starts with the keyword **Sub** and ends with an **End Sub** statement. Here’s an example:

```vb showLineNumbers
Sub Message()
   MsgBox "That’s all folks!"
End Sub
```

This example shows a procedure named `Message`. A set of *parentheses* follows the procedure’s name. 

In most cases, these *parentheses* are empty. However, you may pass arguments to `Sub procedures` from other procedures. 

If your `Sub` uses arguments, list them between the parentheses.

> Please remember when you record a macro with the macro recorder; the result is always a Sub procedure.

## Function procedures

Every `Function` procedure starts with the keyword **Function** and ends with an **End Function** statement. 

Here’s an example:

```vb showLineNumbers
Function CubeRoot(number)
   CubeRoot = number ^ (1/3)
End Function
```
This function, named `CubeRoot`, takes one argument (named `number`), which is enclosed in *parentheses*. 

`Functions` can have any number of arguments or none at all. 

When you execute the function, it returns a *single value* — the cube root of the argument passed to the function.

`VBA` allows you to specify what type of information (also known as `data type`) is returned by a `Function procedure`. 

You can execute a `Function procedure` from another procedure (a `Sub` or another `Function procedure`).

> Please note that no matter how hard you try, you can’t use the macro recorder to record a Function procedure. You must manually enter every Function procedure that you create.

## Naming Subs and Functions

Like humans & pets, every `Sub` and `Function` procedure must have a name. 

Although it is perfectly acceptable to name anything what you want, it’s usually not a good idea to use such a *freewheeling* attitude when `naming procedures`. 

When naming `procedures`, you must follow a **few rules**:

* You can use *letters, numbers, and some punctuation characters*, but the **first** character must be a **letter**.

* You **can’t** use any spaces or periods in the name.

* `VBA` does not distinguish between *uppercase* and *lowercase* letters.

* You **can’t** embed any of the following characters in a procedure name: **#, $, %, &, @, ^, *, or !**

* `Procedure` names can be no longer than *255 characters*. (Of course, you would never make a procedure name this long.)

Ideally, a procedure’s name describes the routine’s purpose.

Some programmers prefer using `sentence-like` names that provide a complete description of the procedure. 

Some examples include `WriteReportToTextFile` and `Get_Print_Options_and_Print_Report`. 

The use of such lengthy names has *pros* and *cons*. 

On the one hand, such names are *descriptive* and usually *unambiguous*. 
On the other hand, they take longer to type. 

Everyone develops a naming style, but the main objectives are to make the names descriptive and to avoid meaningless names such as **Update, Fix, and Macro1**.

Next post will be about `Executing Sub & Function procedures`.
