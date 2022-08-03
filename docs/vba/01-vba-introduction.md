---
title: VBA Introduction
tags:   [VBA]
permalink: /vba/vba-Introduction/
---

To understand `Visual Basic for Application`, lets look at the following *questions*.

## What is Visual Basic for Application?

`Visual Basic for Applications` also known as `VBA`, is a *programming language* developed by **Microsoft**. 

*SOLIDWORKS*, along with *Excel* and other software of *Office Suit*, includes `VBA` language (at no extra cost). 

In a nutshell, `VBA` is a tool that people use to develop program that control *SOLIDWORKS*.

Think about `a robot` that knows everything about *SOLIDWORKS*. This robot can `read instruction`, and it can also `operate` *SOLIDWORKS* very fast and accurate. 

When you want the robot to do something in *SOLIDWORKS*, you write up a set of robotic instruction by using special codes. 

Tell the robot to follow your instruction, while you sit back and take some rest. That’s kind of what `VBA` is all about.

## What can you do with VBA?

You know that people use different *CAD softwares*, not just *SOLIDWORKS*, for different tasks. 

Below is a list of some common tasks:

* Creating sketches

* Creating 3D models

* Creating Simple to Complex assemblies of 3D models

* Linking files with external softwares mostly excel and many more tasks

This list could go on and on, but you get the idea. 

My point is simply that a CAD Software like *SOLIDWORKS* used for wide variety of things. 

Everyone visiting this blog has different *needs and expectations*. 

One thing virtually every visitor has common is the need to automate some aspect of their work. That is what VBA is all about.

## What are the advantages and disadvantages of VBA?

In this section, I briefly describe the good things about `VBA` – and I also explore its darker side.

### VBA Advantages

You can `automate` almost anything you do in *SOLIDWORKS*. 

To do so, you write `instructions` that *SOLIDWORKS* carries out. 

Automating a task by using `VBA` offers several advantages:

* *SOLIDWORKS* always executes the tasks in exactly the `same way`. (In most cases consistency is good.)

* *SOLIDWORKS* performs the task much `faster` than you do it manually.

* If you are a good macro programmer, *SOLIDWORKS* `always` performs the task without `error`.

* If you set things properly, someone who don’t know anything about *SOLIDWORKS* can perform task.

* For long, time-consuming tasks, you don’t have to sit in front of your computer and get board. *SOLIDWORKS* does work, while you drink water.

### VBA disadvantages

It’s only fair that I give equal time to listing the `disadvantages` (or potential disadvantages) of `VBA`:

* You have to know how to write programs in `VBA` (but that’s why you are here, right?). Fortunately, it’s not as difficult as you might expect.

* Sometimes, things go wrong. In other words, you can’t blindly assume that your VBA program will always work correctly under all circumstances. Welcome to world of `debugging` and, if others are using your macros, be prepared for `technical support`.

## VBA in nutshell

Just to let you know what you are in for, I’ve prepared a quick summary of what `VBA` is all about.

* You perform actions in `VBA` by writing (or recording) *code* in a **VBA module**. You view and edit **VBA modules** by using the `Visual Basic Editor (VBE)`.

* A **VBA module** consists of `Sub procedures`. A `sub procedure` is a chunk of computer code that performs some action on or with objects. The following example shows a simple `Sub procedure` called AddThem. 

This amazing program displays the result of 1 plus 1.

```vb showLineNumbers
Sub AddThem()
    Sum = 1 + 1
    MsgBox ("The answer is " & Sum)
End Sub
```

Next post will be about `Visual Basic Editor` or `VBE`. 
