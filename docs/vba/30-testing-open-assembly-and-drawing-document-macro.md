---
title:  VBA Userforms - Testing Open new Assembly and Drawing document
tags:   [VBA Macro Testing]
permalink: /vba/testing-open-assembly-and-drawing-document-macro/
---

## Introduction

In this post, we **test** our **Open new Assembly and Drawing document** code sample.

This post is supplement of **[VBA Userforms - Open new Assembly and Drawing document](/vba/open-assembly-and-drawing-from-userform/)** post.

> *Please visit above post before this post.*

From **[VBA Userforms - Open new Assembly and Drawing document](/vba/open-assembly-and-drawing-from-userform/)** post we expect following results :

1. **Open Assembly** document when we select "*Assembly document*".

2. **Open Drawing** document when we select "*Drawing document*".

When we run our **VBA macro** we get the expected result.

*Now, as a developer, we want to give a thoroughly tested macro/application to our users.*

For testing our VBA macro, we apply some checks so that macro perform same at all machine.

---

## Code block to check

Below is code block where we want to apply our check.

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

---

## Apply check

In above code there is only **one** check we apply.

***We need to check if get the template value or not.***

For this, we need to put an `If` condition before opening a new document.

Please see below code for condition.

``` vb showLineNumbers
' Checking if we got template path
If defaultTemplate = vbNullString Then
  ' If template path is empty then show message and exit from procedure.
  MsgBox "Failed to open " + DocumentTypeComboBox.Value + " template."
  Exit Sub
End If
```

In above code, we check *if got the template path or not*.

If template path is **empty** then 

1. we *show a message* to user as show in below image.
2. we **end** our `sub` procedure from here.

![error-message-on-empty-template](/assets/vba-images/Open_assembly_and_drawing_from_Userform/error-message-on-empty-template.png)

After adding our check, procedure has following code.

```vb showLineNumbers
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
  
  ' Checking if we got template path
  If defaultTemplate = vbNullString Then
    ' If template path is empty then show message and exit from procedure.
    MsgBox "Failed to open " + DocumentTypeComboBox.Value + " template."
    Exit Sub
  End If

  ' Setting Solidworks document to new Assembly document
  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
  
  ' Hiding the Window aft er opening the selected document
  OpenDocumentWindow.Hide
  
  ' Reset the Index of Combo Box to "0" again
  DocumentTypeComboBox.ListIndex = 0
    
End Sub
```

---

## Cause of Error

*You might wondering how can we have such error?*

We have this error, if the value of templates path is **not set** in option setting.

For reference please see below image.

![default-template-options](/assets/vba-images/Open_assembly_and_drawing_from_Userform/default-template-options.png)

As you can see, in my machine these value are already set.

But they are empty in case of fresh installation.

Hence if someone runs this macro on fresh SOLIDWORKS copy, they might get error message which we give.

---

**This is it !!!**

*I hope this be will helpful to someone!*

If you found anything to **add or update**, please let me know on my *e-mail* which is given in bottom.

this post helps us to **test** our **Open new Assembly and Drawing document** macro.

**We will see this type of testing of all our macros which we in this website.**

*If you like the post then please share it with your friends also.*

*Do let me know by you like this post or not!*

*Till then, Happy learning!!!*
