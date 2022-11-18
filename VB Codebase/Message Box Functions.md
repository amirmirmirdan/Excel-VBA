A MsgBox is nothing but a dialog box that you can use to inform your users by showing a custom message or get some basic inputs (such as Yes/No or OK/Cancel).

While the MsgBox dialog box is displayed, your VBA code is halted. You need to click any of the buttons in the MsgBox to run the remaining VBA code.

## Anatomy of a VBA MsgBox in Excel

A message box has the following parts:

1.  **Title**: This is typically used to display what the message box is about. If you don’t specify anything, it displays the application name – which is Microsoft Excel in this case.
2.  **Prompt**: This is the message that you want to display. You can use this space to write a couple of lines or even display tables/data here.
3.  **Button(s)**: While OK is the default button, you can customize it to show buttons such as Yes/No, Yes/No/Cancel, Retry/Ignore, etc.
4.  **Close Icon**: You can close the message box by clicking on the close icon.

## Syntax of the VBA MsgBox Function

As I mentioned, MsgBox is a function and has a syntax similar to other VBA functions.

**_MsgBox( prompt \[, buttons \] \[, title \] \[, helpfile, context \] )_**

-   **prompt** – This is a required argument. It displays the message that you see in the MsgBox. In our example, the text “This is a sample MsgBox” is the ‘prompt’. You can use up to 1024 characters in the prompt, and can also use it to display the values of variables. In case you want to show a prompt that has multiple lines, you can do that as well (more on this later in this tutorial).
-   **\[buttons\]** – It determines what buttons and icons are displayed in the MsgBox. For example, if I use vbOkOnly, it will show only the OK button, and if I use vbOKCancel, it will show both the OK and Cancel buttons. I will cover different kinds of buttons later in this tutorial.
-   **\[title\]** – Here you can specify what caption you want in the message dialog box. This is displayed in the title bar of the MsgBox. If you don’t specify anything, it will show the name of the application.
-   **\[helpfile\]** – You can specify a help file that can be accessed when a user clicks on the Help button. The help button would appear only when you use the button code for it. If you’re using a help file, you also need to also specify the context argument.
-   **\[context\]** – It is a numeric expression that is the Help context number assigned to the appropriate Help topic.

If you’re new to the concept of Msgbox, feel free to ignore the \[helpfile\] and \[context\] arguments. I have rarely seen these being used.

Note: All the arguments in square brackets are optional. Only the ‘prompt’ argument is mandatory.

## Excel VBA MsgBox Button Constants (Examples)

There are different types of buttons that can be used with a VBA MsgBox.


<table><tbody><tr><td><strong>Button Constant</strong></td><td><strong>Description</strong></td></tr><tr><td>vbOKOnly</td><td>Shows only the OK button</td></tr><tr><td>vbOKCancel</td><td>Shows the OK and Cancel buttons</td></tr><tr><td>vbAbortRetryIgnore</td><td>Shows the Abort, Retry, and Ignore buttons</td></tr><tr><td>vbYesNo</td><td>Shows the Yes and No buttons</td></tr><tr><td>vbYesNoCancel</td><td>Shows the Yes, No, and Cancel buttons</td></tr><tr><td>vbRetryCancel</td><td>Shows the Retry and Cancel buttons</td></tr><tr><td>vbMsgBoxHelpButton</td><td>Shows the Help button. For this to work, you need to use the help and context arguments in the MsgBox function</td></tr><tr><td>vbDefaultButton1</td><td>Makes the first button default. You can change the number to change the default button. For example, vbDefaultButton2 makes the second button as the default</td></tr></tbody></table>

## Excel VBA MsgBox Icon Constant

<table><tbody><tr><td><strong>Icon Constant</strong></td><td><strong>Description</strong></td></tr><tr><td>vbCritical</td><td>Shows the critical message icon</td></tr><tr><td>vbQuestion</td><td>Shows the question icon</td></tr><tr><td>vbExclamation</td><td>Shows the warning message icon</td></tr><tr><td>vbInformation</td><td>Shows the information icon</td></tr></tbody></table>

## Assign MsgBox Value to a Variable
Below is a table that shows the exact values and the constant returned by the MsgBox function. You don’t need to memorize these, just be aware of it and you can use the constants which are easier to use.

<table><tbody><tr><td><strong>Button Clicked</strong></td><td><strong>Constant</strong></td><td><strong>Value</strong></td></tr><tr><td>Ok</td><td>vbOk</td><td>1</td></tr><tr><td>Cancel</td><td>vbCancel</td><td>2</td></tr><tr><td>Abort</td><td>vbAbort</td><td>3</td></tr><tr><td>Retry</td><td>vbRetry</td><td>4</td></tr><tr><td>Ignore</td><td>vbIgnore</td><td>5</td></tr><tr><td>Yes</td><td>vbYes</td><td>6</td></tr><tr><td>No</td><td>vbNo</td><td>7</td></tr></tbody></table>

---


## Example
Using the vbYesNo Constant.
```vb
Sub MsgBoxYesNo()
MsgBox "Should we stop?", vbYesNo
End Sub
```

Using the message box as a function to return a boolean, depending on what the user clicked. 
```vb
Function UserConfirmation(StrTitle As String, StrPrompt As String) As Boolean
Dim IntAnswer As Integer
    IntAnswer = MsgBox(StrPrompt, vbYesNoCancel + vbQuestion, StrTitle)
    
        If IntAnswer = vbYes Then ' The variable is equal to 6
            UserConfirmation = True
        Else
            UserConfirmation = False
        End If

End Function

```

^7d82bc

