---
title: VBA Dialog Boxes
tags:   [VBA]
permalink: /vba/dialog-boxes/
---

You can’t use VBA very long without being exposed to dialog boxes. 

They seem to pop up almost every time you select a command. 

VBA uses dialog boxes to obtain information, clarify commands, and display messages. 

If you develop VBA macros, you can create your own dialog boxes that work just like those built in. 

Those custom dialog boxes are called `UserForms` in VBA. About which we look into next section.

## UserForm Alternatives

Some of the VBA macros you create behave the same every time you execute them. 

For example, you may develop a macro for intermediate steps you do every day. 

This macro always produces the same result and requires no additional user input.

You might develop other macros that behave differently under various circumstances or that offer the user options. 

In such cases, the macro may benefit from a custom dialog box. 

A custom dialog box provides a simple means for getting information from the user. 

Your macro then uses that information to determine what it should do.

`UserForms` can be quite useful, but creating them takes time. 

Before I cover the topic of creating UserForms in the next section, you need to know about some potentially timesaving alternatives.

VBA lets you display several different types of dialog boxes that you can sometimes use in place of a `UserForm`. 

You can customize these built-in dialog boxes in some ways, but they certainly don’t offer the options available in a UserForm. 

In some cases, however, they’re just what you need.

In the following sections you read about

* VBA `MsgBox` function

* VBA `InputBox` function

* VBA `GetOpenFilename` method

* VBA `GetSaveAsFilename` method

* VBA `FileDialog` method

Next post will be about ***VBA MsgBox Function***.

