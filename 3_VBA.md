# VBA Guide
VBA is a tool for programming, editing and running application code. You can create a custom function with VBA. 

## Object Oriented Programming Languages
- In each statement start by defining the object. Think of a cell as the object you want to change.
- The basic of OOP Objects are data structures that contain data about an object as well in the form of properties and also code in the form of methods.

## Turning on the Developer Tab
- Go to the file tab, click options, click customize ribbons, check the developer box.
- Use Record Macro, title the macro name and fill descriptions, then save in your personal macro workbook to use on any file or save on 'this workbook' to keep it only in that file.
> TIP: Use 'ctrl + shift + YourKey' as a general formula for saving shortcuts. There are not a lot of shortcuts using this combo.
- Use ctrl + enter to execute each macro step, this creates an open source pathway. Then stop recording.
- Each macro defaults to absolute references. 'Use relative reference' button to create a macro that doesn't depend on an absolute reference. Absolute referencing is good for sorting and filtering. Always consider wether you want to use relative reference of not. 

## Simple Macros with Record Macro 
- Create a few simple macros, like convert a column that was inputted into ecxcel into the correct format. In this case, record the macro and do what you normally would and it will save it for you.
- Get comforatble writing down every step of the macro. 
- The 'insert' button allows you to create a button for your macro.
- To create a protect worksheet macro, click home, find and select tab, Go to special, click constants (only numbers). You've now locked all input numbers.

## Scripting Macros
- Click the Visual Basic app in the Developer tab.
- On the left hand side is the project explorer and the properties window. The gray area contains the code window, where you write vba code.
- The project explorer contains all macros you've recorded in that file called module files. Click the Insert tab, click module to create a folder of all modules.
- Click view to see the immediate window, allows you to ask questions of your workbook and run bits of code.
- Click view, object browser. Object browser is a library of all objects, collections, properties, and methods. Good reference.


## Authoring Macros
- All statement begin with 'Sub' TitleOfMacro() and end with 'End Sub'.
- Apostraphes are used to comment out. Add a description for keyboard shortcut. Comment out every line you're thinking about deleting.
- Use the play button on the top of the coding window to run the sub procedure.
- Use windows key + left arrow to snap windows into the screen.

## Debugging Macros
- Right click on the empty space in the tab of the play icon, and click customize. In the debugging category, click and drag these into the tab; step into, compile project, and toggle breakpoint.
- Excel will not undo macros. Create backups of the original data to test your macro on. 
- To debug, use step into button to run your code line by line and correct output errors. 

## Creating Macros from Scratch
- Objects are stated first in OOP, then you state what you want to do to it using a method (Insert).
- Methods are command statements to do something to the object (verbs). 
- Objects are sheets, tables, charts, cells, columns, and rows.
- Properties are the way we describe certain objects, such as .color, .value, .font (adjectives).

## Macro Scripting Basics
- Go to insert, module, and insert a procedure.
- Procedures are large macros. Sub proecdures are regular macros. Public macros are available to other macros and users. 
Objects, method, 
Rows("1:1").Insert
Range("A:1").Value = "Emp ID"
- There are multiple ways to express the same object. For example, range("C3") is Cells(3,3), is [C3], and use a variable.

## Difference between select and selection
- Select means you selected the object you specified.
- Selection referes to the object you already selected. usually followed by commands. Value = 'YourInput'
- Value is a property.




<details>
 <summary>🛑 SOURCES</summary>

---  
- VBA Beginner Tutorial - https://www.youtube.com/watch?v=G05TrN7nt6k&list=PLoyECfvEFOjYYy54Wa9E83xycKilVMoHp
- 

<ins>Testing</ins> -- To underline text

---

<details>
