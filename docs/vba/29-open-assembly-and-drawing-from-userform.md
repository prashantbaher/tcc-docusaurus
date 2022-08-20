---
title:  VBA Userforms - Open new Assembly and Drawing document
tags:   [VBA]
permalink: /vba/open-assembly-and-drawing-from-userform/
---

In this post, we learn how can we **Open new Assembly and Drawing document** from a Visual Basic for Application's **Userform**.

For this please we need to do following steps as described below.

## Video of Code on YouTube

Please see below video how visually we *Open new Assembly and Drawing document* in **Solidworks VBA macro Userform**.

<iframe src="https://www.youtube.com/embed/AQ3Fyw78ExI" frameborder="0" allowfullscreen></iframe>

Please note that there are **no explanation** in the video. 

**Explanation** of each line and why we write code this way is given in this post.

## Create a new macro

1st we need to create a **new macro** in *Solidworks 3D CAD Software*.

If you don't know how to create a new macro in Solidworks, please go to [VBA in Solidworks](/solidworks-macros/vba-in-solidworks/) post for this.

This will open a new macro in Visual Basic Editor with some code as shown in below image.

![open-new-macro](/assets/vba-images/Open_assembly_and_drawing_from_Userform/open-new-macro.png) 

## Insert userform in the macro

After this we need to insert *a userform* in our macro.

For this, select the button shown in below image.

![insert-userform-into-macro](/assets/vba-images/Open_assembly_and_drawing_from_Userform/insert-userform-into-macro.png)

This button is called ***insert userform***. 

As the name suggest, function of this button is *inserting a userform*.

> Please note that in a macro we can insert any number of userform as we like. But for this example we insert only 1 userform.

After clicking the ***insert userform*** button we get the userform window.

## Adding Controls into Userform

Now in our userform window, we add following controls:

1. **A ComboBox**

2. **A CommandButton**

### Adding ComboBox

You can find `ComboBox` option, as highlighted in *Red Square* in below image.

![insert-combox-into-userform](/assets/vba-images/Open_assembly_and_drawing_from_Userform/insert-combox-into-userform.png)

After adding ComboBox, we get window as shown in below image.

![combobox-into-userform](/assets/vba-images/Open_assembly_and_drawing_from_Userform/combobox-into-userform.png)

### Adding CommandButton

You can find `CommandButton` option, as highlighted in *Red Square* in below image.

![insert-command-button-into-userform](/assets/vba-images/Open_assembly_and_drawing_from_Userform/insert-command-button-into-userform.png)

After adding CommandButton, we get window as shown in below image.

![command-button-into-userform](/assets/vba-images/Open_assembly_and_drawing_from_Userform/command-button-into-userform.png)

## Updating Properties

Now we update some properties of following:

1. **UserForm**

2. **ComboBox**

3. **CommandButton**

### Updating Properties of the UserForm

We update following properties of the Userform:

1. Name of Userform

2. Caption of Userform

In below image, I have shown the properties of `Userform1` and update the properties:

![update-userform-properties](/assets/vba-images/Open_assembly_and_drawing_from_Userform/update-userform-properties.png)

Update the value of *Name* property from `UserForm1` to `OpenDocumentWindow`.

- From *Name* property, we access the Userform.

Update the value of *Caption* property from `UserForm1` to `Open Document`.

- From *Caption* property, we update the text appears in the window of our Userform.

> Please note that it is **not necessary** to update properties but it is a good habit to update them for our purpose. 

### Updating Properties of the ComboBox

Now, we update following property of the ComboBox:

1. Name of ComboBox

In below image, I have shown the properties of `ComboBox1` and update the properties:

![update-combobox-properties](/assets/vba-images/Open_assembly_and_drawing_from_Userform/update-combobox-properties.png)

Update the value of *Name* property from `ComboBox1` to `DocumentTypeComboBox`.

- From *Name* property, we access the ComboBox.

### Updating Properties of the Command Button

Now, we update following property of the Command Button:

1. Name of Command Button

2. Caption of Command Button

In below image, I have shown the properties of `CommandButton1` and update the properties:

![update-command-button-properties](/assets/vba-images/Open_assembly_and_drawing_from_Userform/update-command-button-properties.png)

Update the value of *Name* property from `CommandButton1` to `OpenDocumentButton`.

> From *Name* property, we access the Command Button.

Update the value of *Caption* property from `CommandButton1` to `Open Selected Document`.

> From *Caption* property, we update the text appears in the Command Button of our Userform.

## Calling UserForm in Main Module

Now, we need to call the our *Userform* inside main ***module***.

For this goto main `Sub procedure` inside the **main Module**.

Code inside the main Module is as given below.

{% highlight vb showLineNumbers %}
Dim swApp As Object
Sub main()

Set swApp = Application.SldWorks
End Sub
{% endhighlight %}

To call our `Userform`, replace above code with below code:

{% highlight vb showLineNumbers %}
' Main function of our VBA program
Sub main()
  ' Calling our window to show
  OpenDocumentWindow.Show
End Sub
{% endhighlight %}

Above function call our window to appears on screen.

When the window appears on screen, we *1st* select the document we want to *open* and *then* hit the button to open selected document.

## Adding Document list

Before anything we need to *add* a *list* of documents inside our combobox.

From this list, we select the document which we want to open.

In this post, we are listing only "**Assembly**" and "**Drawing**" documents.

For "**Part**" document, I will give you a simple exercise.

To add documents list inside our combobox, we 1st need to go in **Code Window** of userform.

For this, we need to **right click** on userform and select "**View Code**" option.

Please see below image for "*how to open code window of userform*".

![select-view-code-option](/assets/vba-images/Open_assembly_and_drawing_from_Userform/select-view-code-option.png)

After selecting "**View Code**" a *Code window* appears, which is shown in below image.

![behind-userform-code-window](/assets/vba-images/Open_assembly_and_drawing_from_Userform/behind-userform-code-window.png)

Now *before* adding document list we need to understand **one important thing**!!!

We want to add document list, when we **load** our `Userform`.

For this we need to create an `Initialize Function`.

Below `Code` is that `Initialize Function` which load document lists in our combobox.

```vb showLineNumbers
' Call when user load
Private Sub UserForm_Initialize()
  ' Adding items in Combo Box and also set index to '0'
  ' '0' index means by default we can see "Assembly Document" as already filled
  With DocumentTypeComboBox
    .AddItem "Assembly Document"    ' Adding Assembly Document in Combo Box
    .AddItem "Drawing Document"     ' Adding Drawing Document in Combo Box
    .ListIndex = 0                  ' Select list index for default value to show in combobox
  End With
End Sub
```

Now, above code is ***fully commented*** and ***self explanatory*** hence I will not explain it.

## First Test of Macro

After writing code sample in previous section, we will test if every thing is working correct?

By working correct means:

1. When we `Run` macro, is **Open Document** window appears or not?

2. If window appears, then combobox has documents listed in it?

3. If both item listed, then "**Assembly document**" is showing as pre-filled value or not?

4. Since we have not given any functionality to "**Open Document**" button, hence it should not any done anything when clicking!!!

For testing all the above points we need to `Run` the macro as shown in below image.

![first-test-of-macro](/assets/vba-images/Open_assembly_and_drawing_from_Userform/first-test-of-macro.png)

After running the window we got a window as shown in below image.

![sample-userform-window](/assets/vba-images/Open_assembly_and_drawing_from_Userform/sample-userform-window.png)

Please ***do check*** if your macro is running perfectly till now or not!

If not, then I suggest you to read this article again.

## Add Functionality to Open NewPart Button

To add functionality in our `Open Selected Button`, just double click the `Open NewPart Button`.

This will add give some code behind the designer and opens the **code window** of Userform designer.

```vb showLineNumbers
Private Sub OpenDocumentButton_Click()

End Sub
```

We need to update this code for opening new part after clicking the button.

For this replace all above code with below code.

```vb showLineNumbers
Option Explicit

' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2

' Private function of Open New Part Button 
Private Sub OpenDocumentButton_Click()

  ' Setting Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  
  ' Creating string type variable for storing default Assembly location
  Dim defaultTemplate As String
  
  If DocumentTypeComboBox.Value = "Assembly Document" Then
    ' Setting value of this string type variable to "Default Assembly template"
    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)
  Else
    ' Setting value of this string type variable to "Default drawing template" without define paper size
    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)
  End If

  ' Setting Solidworks document to new Assembly document
  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
  
  ' Hiding the Window after opening the selected document
  OpenDocumentWindow.Hide
  
  ' Reset the Index of Combo Box to "0" again
  DocumentTypeComboBox.ListIndex = 0
    
End Sub
```

Now I have added codes in **2 parts**.

In **1st part** I added below code lines *at top* of the code window.

```vb showLineNumbers
Option Explicit

' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
```

Please see below image for more reference.

![add-code-at-top-of-userform-codewindow](/assets/vba-images/Open_assembly_and_drawing_from_Userform/add-code-at-top-of-userform-codewindow.png)

In **2nd part** I added `OpenDocumentButton_Click` function in the code window as shown in below image.

![code-when-button-clicked](/assets/vba-images/Open_assembly_and_drawing_from_Userform/code-when-button-clicked.png)

I have already explained code inside `OpenDocumentButton_Click` function in [Open Assembly and Drawing document](/solidworks-macros/open-assembly-and-drawing).

But here, I have added a condition which is shown in Red colored box in below image.

![condition-that-control-opening-document](/assets/vba-images/Open_assembly_and_drawing_from_Userform/condition-that-control-opening-document.png)

Basically, this condition stated that, if we select "*Assembly Document*" in combobox, then by clicking "*Open Select Document*" button our macro open "**Assembly**" document in *Solidworks*.

Otherwise, it will always open "**Drawing**" document in *Solidworks*.

The above code will *open New part document* when we click the button.

`Run` the macro, and check wheather our macro is working correct document or not!!

If not, send your macro and I will guide you in doing it correctly! 

## Exercise to do

For those who wants to do more I have an exercise!!

1. Add "**Part Document**" in ComboBox list.

2. Make "**Part Document**" as *pre-filled* value in ComboBox.

3. Change the **conditional statement**, so that it can handle **all 3 conditions**!!

Send me the Code of macro in my below **e-mail** and I will ***verify*** it. 

That's it for now.

Hope you learn some use of Userforms and by this post you can get the idea how they works.

I will provide more tutorials on using of Userform time to time.

***Till then do come to visit this blog and Happy learning!!!***

