# VBA-Common-Library
A library of common Excel VBA functions that I have found useful over the years.

<br> 

## How to Use
1. Create a module in excel called `common`. 
2. Copy the functions you need into the `common` module.
3. Call functions by using `common.functionName`.

<br> 

## Base64 Encode
Encodes string to Base64.

**Input**: `String` 
<br> 
**Output**: `String` encoded to base64

<br> 

## Binary To String
Decodes binary to string. _2003 Antonin Foller, http://www.motobit.com_

**Input**: `Variant` as bianary data (ex. VT_UI1 | VT_ARRAY)
<br> 
**Output**: `String`

<br> 

## Convert Range To Delimited Lists
Converts range or named range to delimited lists.

**Input**:
1. `String` Worksheet Name _(ex. "Sheet 1")_
2. `String` Range Name _(ex. "A1:B5" or "clientNames")_
3. `String` Delimiter _(ex. ";" or ", ")_

**Output**: `String` Delimited List (ex. "Hello, World,")

<br> 

## Count Non-Blank Array Items
Count the number of items in an array that contain a value.

**Input**: `Variant` Array _(ex. [1,2, ,4])_
<br> 
**Output**: `Integer` (ex. 3)

<br> 

## Enable Events _(Sub-Routine)_
Enable or disable events and screen updating on the application level.

**Input**: `Boolean` Enable _(ex. True or False)_
<br> 
**Output**: `None`

<br> 

## Find Query In Column
Finds queried value in a specified column and returns the row number where the query is found.

**Required Functions**:
1. `common.getColumnLetter`

**Input**: 
1. `String` Search Worksheet Name _(ex. "Sheet 1")_
2. `String` Search Term _(ex. "foo")_
3. `String` Search Column _(ex. "A:A")_

**Output**: `Integer` Row Number

<br> 

## Find Query In Row
Finds queried value in a specified row and returns the column number where the query is found.

**Required Functions**:
1. `common.getColumnLetter`

**Input**:
1. `String` Search Worksheet Name _(ex. "Sheet 1")_
2. `String` Search Term _(ex. "foo")_
3. `String` Search Row _(ex. "1:1")_

**Output**: `Integer` Column Number

<br> 

## Fuzzy Find _(beta)_
Finds closest match to queried value in specified range.

**Input**:
1. `String` Query _(ex. "foo")_
2. `String` Search Range _(ex. "A1:B5")_
3. `String` Search Sheet Name _(ex. "Sheet 1")_

**Output**: `String` Closest Matched Value

<br> 

## Generate Range of Available Printers _(Sub-Routine)_
Finds specified column header on specified sheet and enters a list of printers available on the network into the column below the header.

**Required Functions**:
1. `common.findQueryInRow`
2. `common.getColumnLetter`

**Input**:
1. `String` Destination Sheet _(ex. "Sheet 1")_
2. `String` Destination Column Header _(ex. "Printer List")_

**Output**: `None`

<br> 

## Get Column Letter
Returns the column letter for specified column number.

**Input**: `Long` Column Number
<br> 
**Output**: `String` Column Letter

<br> 

## Lock All Sheets _(Sub-Routine)_
Locks or unlocks all sheets, unless a sheet is provided then only that sheet will be locked or unlocked.

**Input**:
1. `Boolean` Locked _(ex. `True` or `False`)_
2. `Worksheet` Single Sheet

**Output**: `None`

<br> 

## One Digit Number to Text
Converts any single-digit number to text.

**Input**: `String` Digit _(ex. "5")_
<br> 
**Output**: `String` Text _(ex. "Five")_

<br> 

## Remove Duplicates
Removes duplicate values from specified range.

**Input**:
1. `Worksheet` Origin Worksheet _(ex. `ThisWorkbook.Sheets("Sheet 1")`)_
2. `Range` Origin Range _(ex. `ThisWorkbook.Sheets("Sheet 1").Range("A1:B5")`)_

**Output**: `Scripting.Dictionary` of which the keys are the values from the range with the duplicates removed

<br> 

## Remove Leading String
Removes the specified leading string if it appears at the beginning of the whole string. You can determine the output to be text or an excel formula to achieve the same result.

**Input**:
1. `String` Leading string to be removed _(ex. "foo ")_
2. `String` Whole string _(ex. "foo bar")_
3. `Boolean`  Return Excel formula or text _(ex. `True` or `False`)_

**Output**:
1. `String` Text (ex. "bar")
2. `String`Formula (ex. "IF(LEFT(" & `whole` & ",LEN("& `lead` &"))=""" & `lead` & """,RIGHT(" & `whole` & ",LEN(" & `whole` & ")-LEN(" & `lead` & ")), " & `whole` & ")")

<br> 

## Save Object, Chart, or Shape as Image to Desktop _(Sub-Routine)_
Takes any shape, chart, or other object and exports it to you desktop as a `PNG` file.

**Input**:
1. `String` Name of worksheet where the object resides
2. `String` Name of object
3. `String` File name of exported image

**Output**: `None`

<br> 

## Spell Number as Currency
Takes a number and spells it out in words eg. `1` to "One".

**Required Functions**:
1. `common.oneDigitNumberToText`
2. `common.twoDigitNumberToText`
3. `common.threeDigitNumberToText`

**Input**:
1. `Variant` Number to Spell _(ex. `123.01`)_
2. `String` Name of Currency _(ex. "Dollars")_

**Output**: `String` Text  _(ex. "One Hundred Twenty Three Dollars And One Cent")_

<br> 

## String To Binary
Encodes string to binary. _2003 Antonin Foller, http://www.motobit.com_

**Input**: `String` Text _(ex. "Hello World")_
<br> 
**Output**: `Binary` _(ex. "01001000 01100101 01101100 01101100 01101111 00100000 01010111 01101111 01110010 01101100 01100100")_

<br> 

## Three Digit Number to Text
Converts any three-digit number to text.

**Required Functions**:
1. `common.oneDigitNumberToText`
2. `common.twoDigitNumberToText`

**Input**: `String` Three digit number _(ex. "123")_
<br> 
**Output**: `String` Number as text _(ex. "One Hundred Twenty Three")_

<br> 

## To Camel Case
Converts string to camel case.

**Required Functions**:
1. `common.toPascalCase`

**Input**: `String` Text _(ex. "Hello World")_
<br> 
**Output**: `String` Text _(ex. "helloWorld")_

<br> 

## To Pascal Case
Converts string to pascal case.

**Input**: `String` Text _(ex. "hElLo WoRlD")_
<br> 
**Output**: `String` Text _(ex. "HelloWorld")_

<br> 

## Two Digit Number to Text
Converts any two-digit number to text.

**Required Functions**:
1. `common.oneDigitNumberToText`

**Input**: `String` Two digit number _(ex. "42")_
<br> 
**Output**: `String` Text _(ex. "Forty Two")_