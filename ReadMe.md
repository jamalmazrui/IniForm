# IniForm 

Version 2.0
November 13, 2015
Copyright 2005 - 2015 by Jamal Mazrui
GNU Lesser General Public License (LGPL)
## Contents
- [Description](#description)
- [Installation](#installation)
- [Controls and Layout](#controls-and-layout)
- [Sample input.ini](#sample-input.ini)
- [Control Attributes](#control-attributes)
- [Two Modes of Input and Output](#two-modes-of-input-and-output)
- [Layout Rules](#l-rules)
- [Miscellaneous Points](#miscellaneous-points)
- [Executable and Library Versions](#executable-and-library Versions)
- [Sample Forms](#sample-forms)
- [Development Notes](#development-notes)

## Description

IniForm is a utility to create a Windows dialog, also called a form, based on a file in the .ini format.  The result of user interaction with the form is also stored in a .ini file.  Such forms are thus ini-oriented in both input and output, and this explains the program name, IniForm.  

Nearly any programming language can read and write .ini files, so IniForm makes it possible to create graphical user interface (GUI) forms easily.  A program can either create a form from an existing input.ini file, or create a form by defining input.ini on the fly just before calling the IniForm processor.  After the user closes the form, the program can read the results in output.ini.

## Installation

The IniForm distribution may be installed in any folder, e.g., in `c:\IniForm`.  Although this is not a standard location for programs on a Windows computer, the benefit to a developer is easy navigation to this folder at a command prompt.  Once there, you can readily test different forms by running IniForm.exe with a parameter for the form ddefinition.  

## Controls and Layout

Nine types of controls may appear on a form:  a static label, push button, check box, radio button, single-selection list box, multi-selection list box, single-line edit box, multi-line edit box, and status bar.  These are system controls provided by Microsoft Windows to any program that requests them.  Screen readers and other assistive technology know how to interpret these controls in a friendly and reliable manner, partly because they are supported by the Microsoft Active Accessibility (MSAA) API.

A distinguishing feature of IniForm is that it can dynamically generate the layout of a form based on minimal information about the controls defined in the input.ini file.  You do not have to specify the location or size of each control.  IniForm examines the type of each control, the sequence of controls, and the initial data they contain, if any.  It then applies rules for sizing, alignment, and placement to determine the layout of the form.  

The form is then activated with this layout, and the user makes choices, resulting in the output.ini file when complete.  output.ini can include layout settings as well as data choices.  Thus, this file may be used either to obtain data choices for subsequent operations, or to obtain layout settings for defining another form more specifically.

## Sample input.ini

Let us examine the content of a sample input.ini file as follows:

```
[Customer Information]
control=form
output=all

[First Name]
control=edit
value=John

[Last Name]
value=Doe
align=r

[Receive newsletter]
control=check
align=d
value=1
tip=Use spacebar to toggle selection

[Receive advertizing]
control=check

[Standard shipping]
control=radio
tip=Use arrow keys to choose one

[2 day ground]
control=radio

[Overnight express]
control=radio

[Address Type]
control=list
range=home|work|vacation|other
selection=2
tip=Choose one

[Recreational Interests]
control=multi
range=fishing|tennis|skiing|baseball|running|soccer|basketball|football|golf
selection=2|4|6
Focus=3
tip=Use spacebar to select one or more choices

[Comments]
control=memo
value=These are miscellaneous notes.|This is a new line.
tip=Type many lines if you want

[OK]
control=button

[Cancel]
control=button
```

In a .ini file like the one shown above, the name of a section is enclosed in square brackets and on a line by itself.  Under the section are one or more lines containing an attribute and a value, separated by an equals sign.  In the example, the first section is called "Customer Information."  Within that section is an attribute called "control."  the value of control is "form."  Section and attribute names are not case sensitive.

The original definition of the .ini format had more limititations.  names of sections and keys could not include spaces.  A value could not be longer than 256 characters.  The total size of the file could not be larger than 64K.  These limitations have been removed over the years.  The main requirements now are that square brackets enclose a section name, that an equals sign separate an attribute name from its value, and that a value not contain carriage or line feed characters, so that it occupies a single line.  Note that a long line may wrap in an editing program if "word wrap" is on, so turn off this setting if manually editing a .ini file.

The input.ini file uses a section for each control of the form.  In addition, the initial section defines settings of the overall form container, rather than a particular control within it.  Each section must uniquely name a control.  

In the example, the name of the form is "Customer Information," which will be the title of the dialog window.  The "output=all" setting is explained later.

The next section is "First Name," corresponding to the first control of the form.  The type of this control is "edit," an abbreviation for a single-line edit box.  The initial value of the control is "John."

The next control is "Last Name."  No control type is specified, so the default type of "edit" is assumed.  The "Align=r" line means that this control should be aligned to the right of the previous one.  To help discuss alignment, controls in the same horizontal region from left to right are said to be in the same "band."

The next two controls are check boxes in the second band of the form, indicated by a downward alignment.  Notice the Tip attribute with text that will appear in the status bar at the bottom of the form when the first check box has keyboard focus.  Also notice that this check box is turned on by default.

An alignment attribute of "r" for right, or "d" for down," only needs to be specified to avoid default placement values.  If not otherwise specified, push buttons, check boxes, and radio buttons are assumed to be alighned to the right in the same band.  On the other hand, list boxes (whether single- or multi-selection) and edit boxes (whether single- or multi-line) are assumed to be aligned downward to a new band.

Band 3 contains thre radio buttons.  Consecutive radio buttons are assumed to be part of the same group that is navigable with the left and right arrow keys.

Band 4 contains a single-selection list box.  The Range attribute contains all possible items, which are separated by a vertical bar (|) character.  The Selection attribute indicates that the 2nd item will be selected when the form is activated.  Band 5 contains a sorted multi-selection list box.  Initially, three items will be selected, and focus will be on the 3rd one.

Band 6 contains a multi-line edit box.  A vertical bar indicates a hard carriage return--otherwise, text will wrap at the right margin of the control.

Band 7 contains OK and Cancel buttons.  Band 8 contains the status bar, where tip text is displayed.

## Control Attributes

Here are the possible attributes for each control:

* Caption -- The external or display name of a control if it is different from the internal one that names the section.

* Control -- The type of control.  Abbreviations are label for static label, button for push button, check for check box, radio for radio button, list for single-selection list box, multi for multi-selection list box, edit for single-line edit box, memo for multi-line edit box, and status for status bar.  If no Control value is specified, the default value is "edit."

* ID -- Number that uniquely represents a control in a form.  These numbers are assigned automatically beginning with 100 if not specified.  However, an OK button is assigned 1, and Cancel is assigned 2--based on Windows conventions.

* Left -- Horizontal coordinate of the upper left corner of the control indicating the distance, in Windows dialog units, from the left border of the form.

* Top -- Vertical coordinate of the upper left corner of the control, indicating the vertical distance, in Windows dialog units, from the top border of the form.

* Width -- Horizontal distance, in Windows dialog units, between the left and right borders of the form.

* Height -- Vertical distance, in Windows dialog units, between the top and bottom borders of the form.

* Style -- Number indicating Windows style of the control, computed by combining Windows bit flags.  Default is assigned based on control type.

* Extend -- Number indicating Windows extended style of the control, computed by combining Windows bit flags.  Default is assigned based on control type.

* Value -- Current content of the control

* Range -- Possible values of a control from which one or more may be chosen

* Selection -- Values currently selected within the range of possible values

* Tip -- Text to be displayed in the status bar when the control has focus.

* Help -- Text to be displayed when the user presses F1 when the control has focus.  Use a vertical bar (|) to specify a hard carriage return within the text.

* Misc -- Miscellaneous information not covered by other attributes.  Optional values are "Password" or "ReadOnly" for edit or memo controls, and "Sort" for list or multi controls.  Separate multiple values with the vertical bar character.

## Two Modes of Input and Output

Attributes called "Input" and "Output" only have meaning in the initial section that references the overall form.  Their values may be either "data" or "all."  If "Input=data," Iniform assumes that only minimal data is provided for controls, such as their types and values but not position or placement.  IniForm thus dynamically generates the layout settings.  If "Input=all," on the other hand, IniForm assumes that both data and layout are specified, so it produces the form according to all settings.

If "Output=data," IniForm produces the output.ini file with only a single section called "Results."  In that section, the attributes are the names of the controls, and the values are the ones the user chose before saving the form.  For example, "Last Name=Smith" might be a line indicating that "Smith" was the value in the Last Name edit box when the OK button was pressed.

If no "Input" attribute is specified, the default value is "data."  Similarly, if no "Output" attribute is specified, the default value is "Data."  To produce an output.ini file with exact layout information, specify "Output=all" in the input.ini file.  You could then copy the output.ini file to produce a new input.ini file.  You could tweak some of the data or layout settings in the new input.ini before using it to activate a form.  

Here is an example of the "Results" section of output.ini:

```
[Results]
First Name=John
Last Name=Smith
Receive newsletter=1
Receive advertizing=0
Standard shipping=1
2 day ground=0
Overnight express=0
Address Type=work
Recreational Interests=basketball|football|running
OK=1
```

## Layout Rules

IniForm generates a form layout according to a set of rules for universal design, addressing both visual and nonvisual usability.  Specifically, the following decisions are made:

*  The top, bottom, left, and right borders of the form are separated from controls by at least 7 dialog units.

*  Labels are created for controls that do not have a built-in caption.  Push buttons, check boxes, and radio buttons have captions, but list boxes do not, so  labels are automatically created for them based on their name.  The label is placed to the left of its associated control, followed by a colon character (:), and separated by 4 dialog units from the control.

*  Each control, or control with label, is separated from another control by at least 7 dialog units, both horizontally and vertically.

*  Controls in the same horizontal band of the form are separated by equal amounts.  The same distance also separates the leftmost control from the left border and the rightmost control from the right border.

*  Push buttons, check boxes, and radio buttons in the same band have the same width, which is the width large enough for the control with the longest caption.

*  All auto-created labels have the same width and their text is right-aligned within the control.  This means that, by default, edit boxes on subsequent lines will have the colon characters of their labels directly above one another.

*  List boxes are sized wide enough to accommodate the widest item they contain.

*  The form as a whole is sized with the width and height needed to accommodate all controls.  It is also centered on the screen.
## Miscellaneous Points

You can specify the value "NoLabel" in the Misc attribute of a control with no caption in order to prevent a label from being automatically created.  Similarly, specify "NoStatus" in the Misc attribute of the form (initial section) to prevent a status bar from being automatically created.

You can search for text in an edit, memo, list, or multi control.  For example, press Control+F for a forward find in a memo control, which can contain up to two megabytes of text.  F3 searches again.  Similarly, Control+Shift+F does a reverse find, and Shift+F3 repeats that search.  A list or multi control can contain thousands of items, so searching for a substring may be quicker than other ways of locating items of interest such as initial letter navigation.

You can associate a hot key with a control by placing an ampersand (&) character in its caption.  For example, the line "Caption=&Phone" could make Alt+P move focus to the Phone field of a form.

You can define list boxes or edit boxes by another method either because their content is large or because it includes vertical bar characters where using those characters does not substitute well for carriage return/line feed pairs.  To do so, create a file called input.txt instead of input.ini.  Put the name of the control in doubled square brackets above its multi-lined text (this helps to ensure that the control name will not be mistaken for bracketed text within the body of the control, itself).  For example, the following is an input.txt file with the "range" of the multi control and the "value" of the memo control above:

```
[[Recreational Interests]]
fishing
tennis
skiing
baseball
running
soccer
basketball
football
golf

[[Comments]]
These are miscellaneous notes.
This is a new line.
```

If IniForm does not find a list box range or a memo value in the input.ini file, it will look for it in an input.txt file.  Similarly, help text may be defined either by a Help attribute with vertical bar characters for line breaks or by a help.txt file with the control name in doubled square brackets above the help text.

## Executable and Library Versions

IniForm is available in two 32-bit versions:  an executable (.exe file) and a COM server library (.dll file).  This flexibility allows it to be invoked within applications built with almost any programming language.  Only 32-bit applications can call the 32-bit IniForm COM server.  Either 32 or 64-bit applications can call the 32-bit Iniform executable.

In either case, IniForm takes a single parameter indicating the source, input.ini file to use.  If no parameter is specified, input.ini is assumed in the default folder (explained later).  If another input definition is desired, specify its root name before the suffix of "_input.ini."  For example, if you ran the following:  
`IniForm.exe welcome`

IniForm would look for welcome_input.ini in the default folder, and it would produce welcome_output.ini when done.  If you want to specify an input file in a particular folder, include its path before the root name, e.g.,  
`IniForm.exe c:\temp\welcome`

would load   
`c:\temp\welcome_input.ini`

and produce   
`c:\temp\welcome_output.ini`

It is also possible to specify a folder path, and the file input.ini will be assumed within that folder, e.g.,  
`c:\temp`

would refer to  
`c:\temp\input.ini`

With the executable version of IniForm, the default folder is the one containing the IniForm.exe file.  With the DLL version, the default folder is the one containing the program that calls IniForm.dll.  Specifying the folder within the source parameter removes ambiguity.  Note that, as a reminder, a source file is specified by less than its complete file name, because IniForm works with a set of inter-related files that end in either "input.ini" or "output.ini" depending on whether the information is loaded into or generated by the program.

A script or application can shell out to IniForm.exe, passing it the source form to run.  After the user closes the form, the application can retrieve results from the generated output.ini file.  Depending on circumstances, the application can operate with source files generated at runtime or located in a temporary folder, which may be deleted afterward.

The IniForm.dll file may be placed in any folder, but it must be registered as a COM server.  Installation programs typically include this capability.  Doing so may be manually achieved with the following command:  
`regsvr32 IniForm.dll

Create an IniForm object using "IniForm" as the ProgID of the COM server.  Three possible methods may be called with the (late binding, dispatch type) object created:  RunForm, ShowResults and GetResult.  RunForm takes the source input.ini file as a single parameter and returns True if a corresponding output.ini was generated.  ShowResults displays the output results of running the form in a message box.  GetResult retrieves a particular result as a string by specifying its control caption.  The batch file, RunTestIniform.bat, may be run to demonstrate use of the IniForm COM server in a VBScript program, TestIniForm.vbs.

## Sample Forms

The IniForm program folder contains several sample forms as follows:

* input.ini and help.txt -- The extensive, Customer Information example used in this documentation.  A resulting output.ini is also included with layout generated.

* ChooseButton_input.ini A set of buttons from which to make a choice.

* LogOn_input.ini -- A Log On form that prompts for a user name and password.

* MultiEdit_input.ini -- A set of edit controls for collecting contact data.

* PickItem_input.ini -- A list control for picking a fruit item.

PickItems_input.ini and PickItems_input.txt -- A large multi control for picking lines in this documentation.

* ReadFile_input.ini and ReadFile_input.txt -- A large edit control for reading this documentation.

* WriteMemo_input.ini and WriteMemo_input.txt -- A large edit control for writing multiple lines of text.

## Development Notes

For the technically interested, IniForm is developed with the PowerBASIC compiler from <http://PowerBASIC.com>.

I welcome feedback.  When reporting a problem, the more specifics, the better -- including steps to reproduce the problem if possible.
