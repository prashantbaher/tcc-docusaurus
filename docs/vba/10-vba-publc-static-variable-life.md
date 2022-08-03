---
title: Public, Static and Variable's Life
tags:   [VBA]
permalink: /vba/publc-static-variable-life/
---

In the following sections we will discussed about following topics: 

1. Public Variables

2. Static Variables

3. Life of Variables

Let's look at them one by one.

## Public Variables

If you need to make a variable *available* to all the procedures in all your VBA modules, declare the variable at the module level (in the *Declarations* section) by using the `Public` keyword. 

For example, in previous figure, if I use `Public` in place of `Dim` in declaration section of **VBE**, then you can use those variables in other procedures of same modules, and for other modules also. 

If you would like a variable to be available to other modules, you must declare the variable as `Public`. 

In practice, sharing a variable across modules is hardly ever done. 

But I guess it’s nice to know that it can be done. 

## Static Variables

Normally, when a procedure ends, all the procedure’s variables are reset. 

**Static** variables are a special case because they retain their value even when the procedure ends. 

You declare a static variable at the *procedure level*. 

A static variable may be useful if you need to track the number of times you execute a procedure. 

## Life of Variables

Nothing lives forever, including **variables**. 

The scope of a variable not only determines where that variable may be used, it also affects under which circumstances the variable is removed from memory. 

You can *purge* (remove) all variables from memory by using three methods:

* Click the *Reset* toolbar button (the *little blue* square button on the *Standard* toolbar in previous figure).

* Click `End` when a *runtime error* message dialog box shows up.

* Include an `End` statement anywhere in your code. This is not the same as an `End Sub` or `End Function` statement. Generally it is an Exit statement. 

Otherwise, only procedure-level variables will be removed from memory when the *macro code* has completed running.
 
Static variables, module level variables, and global (`public`) variables all retain their values in between runs of your code. 

> If you use *module-level* or *global-level* variables, make sure they have the value you expect them to have. You never know whether one of the situations I just mentioned may have caused your variables to lose their content! 

Next post will be about ***VBA Constants***.
