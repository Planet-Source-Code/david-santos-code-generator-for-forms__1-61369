FormFinder
----------
A useful utility to automate code generation for controls in Visual Basic.
Writes repetetive code for TextBoxes, ComboBoxes and CommandButtons.

First, load a Form into Formfinder.

Use the option buttons to select the types of controls you want to generate code for.

Properties Tab
==============
Set Button.Enabled to - generates code that sets all selected buttons Enabled properties to the selected value

Put a check in the Text.Locked and Text.Enabled checkboxes to included that property in the code generation.

I use this primarily in database apps where a whole slew of controls need to be cleared out.


Key Verification Tab
====================
Allow only... - Generates code for selected textboxes. If Others is checked, enter the characters that will 
be allowed in the textbox in the box below. Case sensitive, no delimiters.

Allow everything except... - Generates code for selected textboxes. If Others is checked, enter the characters that will 
not be allowed in the textbox in the box below. Case sensitive, no delimiters.

Useful for data verification.

You should only use one of the above.


Text Verification Tab
=====================
Tired of verifying for blanks in textboxes by:

If Trim(text1.text)="" And Trim(text2.text) = "" And ...

This makes things easy.

You can set to check for = or <>.
Usually check for blanks, so leave Value empty
You can verify each item separately

	If Trim(text1.text)="" Then
	...
	End If

	If Trim(text2.text) = "" Then
	...
	End if

or group them with OR or AND

	If Trim(text1.text)="" Or Trim(text2.text) = "" Or ...

	If Trim(text1.text)="" And Trim(text2.text) = "" And ...


If verifying each separately, you can shoose to setfocus if true

	If Trim(text1.text)="" Then
		Text1.SetFocus 
	End If

	If Trim(text2.text) = "" Then
		Text2.SetFocus 
	End if



Focus Tab
=========
Set focus on the next item by pressing enter, or left if the cursor is at the end of the text.
(work in progress)


