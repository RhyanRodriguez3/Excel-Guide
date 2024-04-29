# VBA Guide
Visual Basic for Applications (VBA) is Microsoft's tool for programming and running application code. It's primarily used in Excel to automate manual tasks such as data cleansing and reporting. VBA is a great introduction to Object Object Oriented Programming (OOP). OOP is a type of programming language.
</br>

## Getting started with Recording Macros on the Developer Tab
1. Go to the file tab, click options, click customize ribbons, check the developer box.
1. Practice using Record Macro with manual tasks. Give it a title, description, and save with personal or on this workbook. Personal allows you to use the macro on any file, while 'this workbook' only allows that file.
> TIP: Use ctrl + enter to execute each macro step, this eliminates any unnecessary VB code. Then stop recording.
> 
> TIP: Use 'ctrl + shift + YourKey' as a general formula for saving shortcuts. There are not a lot of shortcuts using this combo.
1. Always consider wether you want to use relative reference of not. Record macro defaults to using absolute references. 'Use relative reference' button to create a macro that isn't an absolute reference. Absolute referencing is good for sorting and filtering. 
</br>

## Beginner Scripting
Statements start by defining the object, then stating what you want to do to the object. Statements are made with syntax, objects, properties, and methods.
1. Syntax are the elements of grammar. For example `.`
1. Objects are the item you want to manipulate. In excel, these are cell(s), row(s), column(s), table(s), chart(s), table(s), and sheet(s). 
   - `ActiveSheet` is the sheet you are currently on
   - `Sheet(1)` refers to the first sheet in your workbook.
   - `Sheets("InputSheetsName").Select` Great for selecting multiple sheets
1. Methods are built-in functions that allow you to do something to the object (verbs). 
1. Properties are object descriptions (adjectives). `color`, `value`, `font`.
- Variables are small pockets of computer memory used to store and retrieve information. Declare variables using `Dim`, stands for Declared In Memory. Then you state the data type of the object your storing, using as DataType. Give it a value with `=`. Diim defaults to variant data type, which is costly in bytes, so it's best practice to set the data type to save memory. 
  - Data type are
      - Text: Variant (all characters), String, and Char.
      - Number: Decimal, Double (holds little decimals), and Integer
      - Boolean: True / False
      - Misc such as Date, Object, and Range.
```
Dim NameYourVariable(aDim) as DataType
aDim = Range("A1").Value 
```
- 'Option Excplicit' is used above the sub procedure that forces users to set the datatype of their variables.

</br>
 




## Scripting Macros
- All statement begin with `'Sub' TitleOfMacro()` and end with `'End Sub'`.
- On the left hand side is the project explorer and the properties window. The gray area contains the code window, where you write vba code.
- To create a protect worksheet macro, click home, find and select tab, Go to special, click constants (only numbers). You've now locked all input numbers.
- The project explorer contains all macros you've recorded in that file called module files. Click the Insert tab, click module to create a folder of all modules.
- Click view to see the immediate window, allows you to ask questions of your workbook and run bits of code.
- Click view, object browser. Object browser is a library of all objects, collections, properties, and methods. Good reference.
- Apostraphes are used to comment out. Add a description for keyboard shortcut. Comment out every line you're thinking about deleting. Get comfortable writing down every step of the macro.
- Use the play button on the top of the coding window to run the sub procedure.
- Use windows key + left arrow to snap windows into the screen.
- Right click on the empty space in the tab of the play icon, and click customize. In the debugging category, click and drag these into the tab; step into, compile project, and toggle breakpoint.
- Excel will not undo macros. Create backups of the original data to test your macro on. 
- To debug, use step into button to run your code line by line and correct output errors. 
- Go to insert, module, and insert a procedure.



## Macro Scripting Basics
- Procedures are large macros. Sub procedures are regular macros. Public macros are available to other macros and users. 
- There are multiple ways to express the same object. For example,
   - range("C3")
   - Cells(3,3)
   - is [C3]



## Difference between select and selection
- Select means you selected the object you specified.
- Selection referes to the object you already selected. Usually followed by commands. Value = 'YourInput'
  - Selection.CurrentRegion
- Properties.
  - .Value aka FormulaR1C1 is the archaic coding language
  - .Name
  - `Insert`
  - .CurrentRegion
 - With and End With come before the object and apply all the commands. Usually prefaced with a '.' The With statement allows you to make an object statement once refer to all the properties of that object.



## Useful Codes
- Click on the current list of data you have and this code will select the entire dataset.
```
Selection.CurrentRegion.Select
```
- You can refer to an object and its current values. The code below inputs the value of the cell as whatever the name of the sheet is then adds your value
```
Range("A1").Value = ActiveSheet.Name & "InputYourValue"
```

<!-- START -->
<details>
 <summary>🤖 Loop Statements</summary>

---
There are different types of loops in VBA, For-Next loops and the Do loop. The steps to create a loop are
1. Loops require variables. FIrst define your variable.
1. Write the code you want to repeat.
1. Wrap that code with a For Next Loop
1. Define the iteration count for the For statement
1. Tell the Next statement to move on tot he next variable value.

For-Next Loop
    For X = 10(YourCondition)
    Cells(x,1).Value = 100
    Next x
    
Double For-Next Loop
    Public Sub MacroName()
    Dim aDimCol as Integer
    Dim bDimRow as Integer
     For aDimCol = 1 to 3
     For aDimRow = 1 to 10 
     Cells(bDimRow,aDimCol).Value = "wow" <!-- For rows 1-10 and cols 1-3 insert this value  -->
     Next bDimRow
     Next aDimCol
    End Sub

Triple For-Next Loop
  Public Sub TripleLoopEx()
    Dim intCol as Integer
    Dim intRow as Integer
    Dim intSheet as Integer
      For intSheet = 3 to 5
      For intCol = 1 to 3
      For intRow = 1 to 5
      Worksheets(intsheet).Cells(intRow, intCol).Value = "wow"
      Next intSheet = 3 to 5
      Next intCol = 1 to 3
      Next intRow = 1 to 5 
  End Sub

For Each Loop


---

</details>
<!-- END -->

<!-- START -->
<details>
 <summary>🤖 Do Loops</summary>

---

---

</details>
<!-- END -->

<details>
 <summary>🛑 SOURCES</summary>

---  
- VBA Beginner Tutorial - https://www.youtube.com/watch?v=G05TrN7nt6k&list=PLoyECfvEFOjYYy54Wa9E83xycKilVMoHp
- 

<ins>Testing</ins> -- To underline text

---

<details>
