# VBA Guide
Visual Basic for Applications (VBA) is Microsoft's tool for programming and running application code. It's primarily used in Excel to automate manual tasks such as data cleansing and reporting. VBA is a great introduction to Object Object Oriented Programming (OOP). OOP is a type of programming language.
- I'm learning how to read beginner to intermediate lines of code.
- Defining a few concepts and understanding how to structure a statement. Identifying the problem.  So I think its best to introduce the ideas then create examples where I explain what it does.
- Making one from scratch has to come with practice.

The steps involved in programming with VBA are
1. First identify the problem then write out the manual way you would do each step. Use record Macro for hints. Define the objects and general methods.
2. Declare your variables. This might need to come with practice, because I need to know how VB selects things.
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
- Use step into button to execute each line of the code. Great for testing codes line by line.



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
 <summary>🤖 For Next Loop Statements: Continue for a specified amount</summary>

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

For Each Loop: Works like a for next loop, but operates on collections. Collections like worksheets
   Dim x As Worksheet
    For Each x In Worksheet
    MsgBox "Found Worksheet: " & x.Name <!-- This looks at each worksheet and creates a message box  -->
   Next x

Exit For Statement: Used to exit a loop early 
Dim x as Integer
For x = 1 to 50
Range("B" & x).Select
 IF Range("B" & x).Value = "Stop" THEN
  Exit For
ElseIf Range("B" & x).Value = "" Then
  Range("B" & x).value = "Info"
End If
Next x


---

</details>
<!-- END -->

<!-- START -->
<details>
 <summary>🤖 Do Loops</summary>

---
- Do loops continue until you specify a criteria to complete.
1. Declare your variable
2. Code the statement to declare the starting number of your variable.
3. Type the code you wanted repreated
4. Wrap with a do while loop
5. Determine the stopping point after the do while statement
6. After the code to repeat, enter in some code to up your variable.

```
Public Sub ExampleDoWhileLoop
Dim x as Integer
x = 1
Do while x < 10
Cells(x,1).Value = 100
x = x + 1
Loop  --> This ends the loop
End Sub
```
- Do while not empty continues until the specified range or condition is not empty. Means setting the value to  <>""
- Do Until loops continue until a test is true. Do Loop Until is another iteration of this.

---

</details>
<!-- END -->



<!-- START -->
<details>
 <summary>🤖 Generating reports: Count and Offset</summary>

---
?Worksheets.Count <!-- Worksheeets is the object. Count is the method -->
Selection.Offset() <!-- Offset allows you to travel rows move (negatives) up and (postiives) down and columns move left right from the current selection -->

- To copy data from one sheet to another. Declare the variables, variables are the objects you want to change (objects and their info). Then select each and use count and offset in a for next loop.
```
Dim x as Integer
Dim sheettitle as string
For x = 1 to worksheets.count - 1
Worksheets(x).select    --> Goes to first sheet
Sheettitle = activesheet.name --> Grabs the sheet title and stores it in dim sheettitle
Worksheets("P&L").Select    --> This selects
Range("A1").Select
Selection.Offset
```

---

</details>
<!-- END -->


<!-- START -->
<details>
 <summary>🤖 Useful concepts for generating reports</summary>

---
- To copy data from one sheet to another. Declare the variables, variables are the objects you want to change (objects and their info). Declaring a variable means storing things in it.
- Then select each and use count and offset in a for next loop.
- You can call other macros you've made. Done by writing the method `Call`
  - Instead of control shift down to find the very bottom cell, you can state: Range("A1000000").Select (This selects the very most bottom cell)
  - Then: Selection.End(xlUp).Select
  - ActiveCell.Offset(2,0).Select <!-- This will go down from the selection -->
- Find and replace in VB is 
```
'Code will find the data and remember where it is
Dim datastart as String
Range("A1").Select
Selection.End(xlDown).Select    <!-- This goes down form the selected range and select the next object. -->
datastart = ActiveCell.address    <!-- The address method remembers the reference location of that cell. -->
Range(datastart).Select
```

---

</details>
<!-- END -->



<!-- START -->
<details>
 <summary>🤖 Messages Box and Input Boxes</summary>

---
- Message boxes are pop ups for the user
  - Create a variable as the message box
  - When creating a variable usually the naming convention is "abbreviated datatype" + "Title"
  - When codes move to the right where you can't see them, you can use a code continuation character.
  - 
- Input boxes are inputs that the user can type in

---

</details>
<!-- END -->



<!-- START -->
<details>
 <summary>🤖 If Else Then & Select Case</summary>

---
- If, Then condition1, elseif condition2, and end the statement with End If
- Elseif is used for a specific input and else if for anything nonspecified. 
- `Select Case` is a more efficient way to write if then statements and ends with `End Select`
- You can create multiple variables, use commas to seperate them.
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
