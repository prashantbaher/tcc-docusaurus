---
title: VBA UserForms
tags:   [VBA]
permalink: /vba/userform/
---

A *UserForm* is useful if your VBA macro needs to get information from a user. 

For example, your macro may have some options that can be specified in a UserForm. 

If only a few pieces of information are required (for example, a *Yes/No* answer or a text *string*), one of the techniques I describe in previous articles may do the job. 

But if you need to obtain more information, you must create a UserForm.

To create a UserForm, you usually take the following general steps:

* Determine how the dialog box will be used and where it will be displayed in your VBA macro.

* Activate the VBE and insert a new UserForm object. A UserForm object holds a single UserForm.

* Add controls to the UserForm. Controls include items such as text boxes, buttons, check boxes, and list boxes.

* Use the Properties window to modify the properties for the controls or for the UserForm itself.

* Write *event-handler* procedures for the controls (for example, a macro that executes when the user clicks a button in the dialog box). These procedures are stored in the Code window for the UserForm object.

* Write a *procedure* (stored in a VBA module) that displays the dialog box to the user.

When you are designing a *UserForm*, you are creating what developers call the **Graphical User Interface (GUI)** to your application. 

Take some time to consider what your form should look like and how your users are likely to want to interact with the elements on the UserForm. 

Try to guide them through the steps they need to take on the form by carefully considering the arrangement and wording of the controls. 

Like most things VBA-related, the more you do it, the easier it gets.

## Userforms Working

Each dialog box that you create is stored in its own UserForm object — one dialog box per UserForm. 

You create and access these UserForms in the Visual Basic Editor.

## Inserting a new UserForm

To insert a UserForm object with the following steps:

1. In the macro, you can insert User form with following 2 ways:

  * From "Menu Bar" -> "UserForm"

  * From “Standard Toolbar” by clicking “Insert UserForm” ![A-new-userform-object](/assets/vba-images/Userforms/3.InsertUserformButtoninstandardToolbar.PNG)

  The VBE insert a new UserForm object with an empty dialog box.

2. If “Property window” is not available in your macro, press `F4` to display “Property window”.

The VBE inserts a new UserForm object, which contains an empty dialog box.

Below figure shows a UserForm — an empty dialog box with some controls in Toolbox.

![Empty-userform-object](/assets/vba-images/Userforms/1.Anewuserformobject.PNG)


## Adding controls to a UserForm

When you activate a UserForm, the VBE displays the Toolbox in a floating window, as shown in above figure. 

You use the tools in the Toolbox to add controls to your UserForm. 

If the Toolbox doesn’t appear when you activate your UserForm, choose **View -> Toolbox**.

To add a control, just click the desired control in the Toolbox and drag it into the dialog box to create the control. 

After you add a control, you can move and resize it by using standard techniques.

Below table indicates the various tools, as well as their capabilities. 

To determine which tool is which, hover your mouse pointer over the control and read the small pop-up description.

<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">ToolBox Control</th>
    </tr>
    <tr>
        <th>Controls</th>
        <th>What it does</tr>
    </tr>
    <tr>
        <td>Label</td>
        <td>Shows text</td>
    </tr>
    <tr>
        <td>TextBox</td>
        <td>Determines which of the file filters the dialog box displays by default.</td>
    </tr>
    <tr>
        <td>ComboBox</td>
        <td>
            Display a drop-down list.
        </td>
    </tr>
    <tr>
        <td>ListBox</td>
        <td>
            Display a list of items.
        </td>
    </tr>
    <tr>
        <td>CheckBox</td>
        <td>Useful for On/off or Yes/No options.</td>
    </tr>
    <tr>
        <td>OptionButton</td>
        <td>Used in groups; allows the user to select one of several options.</td>
    </tr>
    <tr>
        <td>ToggleButoon</td>
        <td>A button that is either on or off.</td>
    </tr>
    <tr>
        <td>Frame </td>
        <td>A container for other control.</td>
    </tr>
    <tr>
        <td>CommandButton</td>
        <td>A clickable button.</td>
    </tr>
    <tr>
        <td>TabStrip</td>
        <td>Display Tabs</td>
    </tr>
    <tr>
        <td>MultiPage</td>
        <td>A tabbed container for other objects.</td>
    </tr>
    <tr>
        <td>ScrollBar</td>
        <td>A draggable bar.</td>
    </tr>
    <tr>
        <td>SpinButton</td>
        <td>A clickable button often used for changing a value.</td>
    </tr>
    <tr>
        <td>Image</td>
        <td>Contains an image</td>
    </tr>
    <tr>
        <td>RefEdit</td>
        <td>Allows the user to select a range.</td>
    </tr>
</table>
-->

|Controls|What it does|
|--- |--- |
|Label|Shows text|
|TextBox|Determines which of the file filters the dialog box displays by default.|
|ComboBox|Display a drop-down list.|
|ListBox|Display a list of items.|
|CheckBox|Useful for On/off or Yes/No options.|
|OptionButton|Used in groups; allows the user to select one of several options.|
|ToggleButoon|A button that is either on or off.|
|Frame|A container for other control.|
|CommandButton|A clickable button.|
|TabStrip|Display Tabs|
|MultiPage|A tabbed container for other objects.|
|ScrollBar|A draggable bar.|
|SpinButton|A clickable button often used for changing a value.|
|Image|Contains an image|
|RefEdit|Allows the user to select a range.|

## Changing properties for a UserForm control

Every control you add to a UserForm has a number of properties that determine how the control looks or behaves. 

In addition, the UserForm itsel also has its own set of properties. 

You can change these properties with the *Properties window*. 

Below figure shows the properties window when a `CommandButton` control is selected:

![Empty-userform-object](/assets/vba-images/Userforms/2.UsethePropertiesWindowstoChangethePropertiesofUserFormControls.PNG)

Properties for controls include the following:

* Name
* Width
* Height
* Value
* Caption

Each control has its own set of properties (although many controls have some common properties). To change a property using the Properties window:

1. Make sure that the correct control is selected in the UserForm.
2. Make sure the Properties window is visible (press `F4` if it’s not).
3. In the Properties window, click on the property that you want to change.
4. Make the change in the right portion of the Properties window.

If you select the **UserForm** itself (not a **control** on the UserForm), you can use the Properties window to adjust UserForm properties

> Some of the UserForm properties serve as default settings for new controls you drag onto the UserForm. For example, if you change the Font property for a UserForm, controls that you add will use that same font. Controls that are already on the UserForm are not affected.

## Viewing the UserForm Code window

Every UserForm object has a Code module that holds the VBA code (*the event-handler procedures*) executed when the user works with the dialog box. 

To view the Code module, press `F7`. 

The *Code window* is empty until you add some procedures. Press `Shift+F7` to return to the dialog box.

Here’s another way to switch between the Code window and the UserForm display: 

- Use the View Code and View Object buttons in the Project window’s title bar. 

- Or right-click the UserForm and choose View Code. 

If you’re viewing code, *double-click* the UserForm name in the Project window to return to the UserForm.

## Showing the UserForm

You display a UserForm by using the UserForm’s `Show` method in a VBA procedure.

The macro that displays the dialog box must be in a VBA module — not in the Code window for the UserForm.

The following procedure displays the dialog box named `UserForm1`:

```vb showlinenumbers showLineNumbers
' Showing the UserForm
Sub ShowDialogBox()
  UserForm.Show
  'Other statements can go here
End Sub
```

When Solidworks displays the dialog box, the `ShowDialogBox` macro halts until the user closes the dialog box. 

Then VBA executes any remaining statements in the procedure. 

Most of the time, you won’t have any more code in the procedure.

## Using information from a UserForm

The VBE provides a name for each control you add to a UserForm. 

The control’s name corresponds to its `Name` property. 

Use this name to refer to a particular control in your code. 

For example, if you add a `CheckBox` control to a UserForm named `UserForm1`, the CheckBox control is named `CheckBox1` by default. 

The following statement makes this control appear with a check mark:

```vb showlinenumbers showLineNumbers
UserForm1.CheckBox1.Value = True
```

Most of the time, you write the code for a UserForm in the UserForm’s code module. 

If that’s the case, you can omit the UserForm object qualifier and write the statement like this:

```vb showlinenumbers showLineNumbers
CheckBox1.Value = True
```

> I recommend that you change the default name the VBE has given to your controls to something more meaningful.

This will sum-up our tutorials on Visual Basic for Application. From now on I will give tutorials on how to use Solidworks commands with the help of VBA Macro.

If you want to know any explaination on any topic related to VBA, please drop a comment and I will try to give it to you. 

### Thank you!!!!

## UPDATE:

I have started VBA UserForm Example in this tutorials lists. 

So if you want to learn how I use these Forms, you can watch them in UserForm Example List Post.

<!-- This is post navigation bar 
<div class="w3-bar w3-margin-top w3-margin-bottom">
    <a href="/visual-basic/vba-other-dialog" class="w3-button w3-rose">&#10094; Previous</a>
    <a href="/visual-basic/open-part-from-userform" class="w3-button w3-rose w3-right">Next &#10095;</a>
</div>
-->
