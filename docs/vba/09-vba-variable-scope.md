---
title: Variable Scope
tags:   [VBA]
permalink: /vba/variable-scope/
---

A *variable’s* scope determines which modules and procedures can use the variable.

## Procedure-only Variables

The *lowest* level of scope for a variable is at the **procedure** level. 

A procedure is either a `Sub` or a `Function` procedure. 

Variables declared with this scope can be used only in the procedure in which they are declared. 

When the procedure ends, the variable no longer exists (it goes to the great big bucket in the sky), and your **CAD** application frees up its memory. 

If you execute the procedure again, the variable comes back to life, but its previous value is lost.

The most common way to declare a *procedure-only* variable is with a `Dim` statement. 

`Dim` doesn’t refer to the mental capacity of the VBA designers. 

Rather, it’s an old programming term that’s short for dimension, which simply means you are setting aside memory for a particular variable. 

You usually place `Dim` statements immediately after the `Sub` or `Function` statement and before the procedure’s code. 

The following example shows some procedure-only variables declared by using `Dim` statements: 

```vb showlinenumbers showLineNumbers
Sub MySub()
  Dim x As Integer
  Dim First As Long
  Dim PartDimension As Single
  Dim myValue
' ...[The procedure’s code goes here]...
End Sub 
```

Notice that the last `Dim` statement in the preceding example doesn’t declare a data type; it declares only the variable itself. The effect is that the variable `MyValue` is a *Variant*.

By the way, you can also declare several variables with a single `Dim` statement, as in the following example: 

```vb showlinenumbers showLineNumbers
Dim x As Integer, y As Integer, z As Integer
Dim First As Long, Last As Double
```

Unlike some languages, **VBA** doesn’t allow you to declare a *group* of variables to be a particular *data type* by separating the variables with **commas**. 

For example, though valid, the following statement does not declare all the variables as *Integers*: 

```vb showlinenumbers showLineNumbers
Dim i, j, k As Integer
```

In this example, only `k` is declared to be an *Integer*; the other variables are declared to be *Variants*. 

If you declare a variable with *procedure-only* scope, other procedures in the same module can use the same variable name, but each instance of the variable is unique to its own procedure.

> In general, variables declared at the *procedure level* are the most efficient because **VBA** frees up the memory they use when the procedure ends.

## Module-only Variables

Sometimes, you want a *variable* to be available to *all procedures* in a module. 

If so, just declare the *variable* (using `Dim` or `Private`) before the module’s first `Sub` or `Function` statement — outside any procedures. 

This is done in the **Declarations** section, at the *beginning* of your module. 

This is also where the `Option Explicit` statement is located. 

Below figure shows how you know when you’re working with the **Declarations** section. 

![Variable-Scope-Example](/assets/vba-images/Programming_Concepts/1.VariableExamples.PNG)

As shown in above image, I want a variable named `swApp`, so that it can available to all procedures in this module. 

Hence I declare this variable in **Declaration** section of **VBE**. 

Next post wil be about ***Public, Static and Variable's Life***.
