---
title: VBA GetOpenFilename, GetSaveAsFilename and Getting a Folder Name
tags:   [VBA]
permalink: /vba/other-dialogs/
---

If your VBA procedure needs to ask the user for a filename, you could use the *InputBox* function. 

An input box usually isn’t the best tool for this job, however, because most users find it difficult to remember paths, backslashes, filenames, and file extensions. 

In other words, it’s far too easy to make a **typographical error** when typing a filename.

For a better solution to this problem, use the **GetOpenFilename** method of the Application object, which ensures that your code gets its hands on a valid filename, including its complete path. 

The *GetOpenFilename* method displays the familiar Open dialog box.

The *GetOpenFilename* method doesn’t actually open the specified file. 

This method simply returns the user-selected filename as a `string`. 

Then you can write code to do whatever you want with the filename.

## Syntax for the GetOpenFilename method

The official syntax for the **GetOpenFilename** method is as follows:

```vb showLineNumbers
' The GetOpenFilename method syntax
object.GetOpenFilename ([fileFilter], [filterIndex], [title],[buttonText], [multiSelect])
```

The GetOpenFilename method takes the optional arguments shown in below Table.

<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">The GetOpenFilename method Arguments</th>
    </tr>
    <tr>
        <th>Arguments</th>
        <th>What it does</tr>
    </tr>
    <tr>
        <td>FileFilter</td>
        <td>
            Determines the types of files that appear in the dialog box (for example, *.TXT). 
            You can specify several different filters from which the user can choose.
        </td>
    </tr>
    <tr>
        <td>FilterIndex</td>
        <td>
            Determines which of the file filters the dialog box displays by default.
        </td>
    </tr>
    <tr>
        <td>Title</td>
        <td>
            Specifies the caption for the dialog box’s title bar.
        </td>
    </tr>
    <tr>
        <td>ButtonText</td>
        <td>
            Ignored
        </td>
    </tr>
    <tr>
        <td>MultiSelect</td>
        <td>If True, the user can select multiple files.</td>
    </tr>
</table>
-->

|Arguments|What it does|
|--- |--- |
|FileFilter|Determines the types of files that appear in the dialog box (for example, *.TXT). 
            You can specify several different filters from which the user can choose.|
|FilterIndex|Determines which of the file filters the dialog box displays by default.|
|Title|Specifies the caption for the dialog box’s title bar.|
|ButtonText|Ignored|
|MultiSelect|If True, the user can select multiple files.|

## A GetOpenFilename example

The `fileFilter` argument determines what appears in the dialog box’s Files of Type drop-down list. 

This argument consists of pairs of file filter strings followed by the wild card file filter specification, with commas separating each part and pair. 

If omitted, this argument defaults to the following:

```vb showLineNumbers
' A GetOpenFilename example
All Files (*.*), *.*
```

Notice that this string consists of two parts:

```vb showLineNumbers
All Files (*.*)
```

and

```vb showLineNumbers
*.*
```

The first part of this string is the text displayed in the Files of Type dropdown list. 

The second part determines which files the dialog box displays. For example, *.* means all files.

The code in the following example brings up a dialog box that asks the user for a filename. 

The procedure defines five file filters. 

Notice that I use the VBA line continuation sequence to set up the Filter variable; doing so helps simplify this rather complicated argument.

```vb showLineNumbers
' A GetOpenFilename example
Sub GetImportFileName()
  Dim Finfo As String
  Dim FilterIndex As Integer
  Dim Title As String
  Dim FileName As Variant

  'Set up list of file filters
  If (IsNumeric)NumberOfSheets Then
  FInfo = "Text Files (*.txt),*.txt," & _
  "Lotus Files (*.prn),*.prn," & _
  "Comma Separated Files (*.csv),*.csv," & _
  "ASCII Files (*.asc),*.asc," & _
  "All Files (*.*),*.*"

  'Display *.* by default
  FilterIndex = 5

  'Set the dialog box caption
  Title = "Select a File to Import"

  'Get the filename
  FileName = Application.GetOpenFilename (FInfo, FilterIndex, Title)

  'Handle return info from dialog box
  If FileName = False Then
    MsgBox "No file was selected."
  Else
    MsgBox "You selected " & FileName
  End If
End Sub
```

Notice that the `FileName` variable is declared as a Variant data type. 

If the user clicks `Cancel`, that variable contains a Boolean value (False). 

Otherwise, FileName is a `string`. Therefore, using a Variant data type handles both possibilities.

## GetSaveAsFilename Method

The *GetSaveAsFilename* method works just like the *GetOpenFilename* method, but it displays the Save As dialog box rather than its Open dialog box. 

The *GetSaveAsFilename* method gets a path and filename from the user but doesn’t do anything with it. 

It’s up to you to write code that actually saves the file.

The syntax for this method follows:

```vb showLineNumbers
' The GetSaveAsFilename method syntax
object.GetSaveAsFilename ([InitialFilename], [FileFilter], [FilterIndex], [Title], [ButtonText])
```

The *GetSaveAsFilename* method takes below arguments, all of which are optional.

<!--
<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">The GetSaveAsFilename method Arguments</th>
    </tr>
    <tr>
        <th>Arguments</th>
        <th>What it does</tr>
    </tr>
    <tr>
        <td>InitialFileName</td>
        <td>Specifies a default filename that appears in the File Name box.</td>
    </tr>
    <tr>
        <td>FileFilter</td>
        <td>
            Determines the types of files that appear in the dialog box (for example, *.TXT). 
            You can specify several different filters from which the user can choose.
        </td>
    </tr>
    <tr>
        <td>FilterIndex</td>
        <td>
            Determines which of the file filters the dialog box displays by default.
        </td>
    </tr>
    <tr>
        <td>Title</td>
        <td>
            Specifies the caption for the dialog box’s title bar.
        </td>
    </tr>
</table>
-->

|Arguments|What it does|
|--- |--- |
|InitialFileName|Specifies a default filename that appears in the File Name box.|
|FileFilter|Determines the types of files that appear in the dialog box (for example, *.TXT). 
            You can specify several different filters from which the user can choose.|
|FilterIndex|Determines which of the file filters the dialog box displays by default.|
|Title|Specifies the caption for the dialog box’s title bar.|

## Getting a Folder Name

Sometimes, you don’t need to get a filename; you just need to get a folder name. 

If that’s the case, the *FileDialog* object is just what the doctor ordered.

The following procedure displays a dialog box that allows the user to select a directory. 

The selected directory name (or “Canceled”) is then displayed by using the `MsgBox` function.

```vb showLineNumbers
' FileDialog example
Sub GetAFolder()
  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = Application.DefaultFilePath & "\"
    .Title = "Please select a location for the backup"
    .Show
    If .SelectedItems.Count = 0 Then
      MsgBox "Canceled"
    Else
      MsgBox .SelectedItems(1)
    End If
  End With
End Sub
```

The *FileDialog* object lets you specify the starting directory by specifying a value for the InitialFileName property. 

In this case, the code uses default file path as the starting directory.

Next post will be about ***VBA UserForms***.
