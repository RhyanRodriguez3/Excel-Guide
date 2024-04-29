# VBA Guide
- Visual Basic for Applications (VBA) is Microsoft's tool for programming, editing, and running application code. It is primarily used in Excel to automate manual tasks.
> such as data dump cleansing and reporting.
>
> - Data structures contain data about an object and its properties.

## Object Oriented Programming (OOP) Intro
VBA is an OOP language. This means statements start by defining the object, then stating what you want to do to the object. Statements are made with objects, properties, syntax, and methods.
1. Syntax are the elements of grammar 
1. Objects are the item you want to manipulate. In excel, these are cell(s), row(s), column(s), table(s), chart(s), table(s), and sheet(s).
   - `ActiveSheet` is the sheet you are currently on
   - Sheet(1) refers to the first sheet in your workbook.
   - Sheets("NameOfYourSheet").select Great for multiple sheets
1. Methods are built-in functions. Methods are command statements to do something to the object (verbs). 
1. Properties are the way we describe certain objects, such as .color, .value, .font (adjectives).

 

## Getting started with Recording Macros on the Developer Tab
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
  - Selection.CurrentRegion
- Properties.
  - .Value aka FormulaR1C1 is the archaic coding language
  - .Name
  - .CurrentRegion
 - With and End With come before the object and apply all the commands. Usually prefaced with a '.' The With statement allows you to make an object statement once refer to all the properties of that object.

## Useful Codes
- Click on the current list of data you have and this code will select the entire dataset.
```
Selection.CurrentRegion.Select
```
- You can refer to an object and its current values. The code below inputs the value of the cell as whatever the name of the sheet is then adds your value
```
Range("InputYourCell").Value = ActiveSheet.Name & "InputYourValue"
```


<details>
 <summary>🛑 SOURCES</summary>

---  
- VBA Beginner Tutorial - https://www.youtube.com/watch?v=G05TrN7nt6k&list=PLoyECfvEFOjYYy54Wa9E83xycKilVMoHp
- 

<ins>Testing</ins> -- To underline text

---

<details>
