# VBA-Common-Library
A library of common Excel VBA functions that I have found useful over the years.

## How to Use
1. Create a module in excel called `common`. 
2. Copy the functions you need into the `common` module.
3. Call functions by using `common.functionName`.

## Base64 Encode
Encodes string to Base64.

**Input**: `String` 
<br>
**Output**: `String` encoded to base64

## Binary To String
Decodes binary to string. _2003 Antonin Foller, http://www.motobit.com_

**Input**: `Variant` as bianary data (ex. VT_UI1 | VT_ARRAY)
<br>
**Output**: `String`

## Convert Range To Delimited Lists
Converts range or named range to delimited lists.

**Input**:
1. `String` Worksheet Name _(ex. "Sheet 1")_
2. `String` Range Name _(ex. "A1:B5" or "clientNames")_
3. `String` Delimiter _(ex. ";" or ", ")_
<br>
**Output**: `String` Delimited List (ex. "Hello, World,")

## Count Non-Blank Array Items
Count the number of items in an array that contain a value.

**Input**: `Variant` Array _(ex. [1,2, ,4])_
<br>
**Output**: `Integer` (ex. 3)

## Enable Events _(Sub-Routine)_
Enable or disable events and screen updating on the application level.

**Input**: `Boolean` Enable (ex. True or False)
<br>
**Output**: `None`

## Find Query In Column
Finds queried value in a specified column and returns the row number where the query is found.

**Input**: 
1. `String` Search Worksheet Name _(ex. "Sheet 1")_
2. `String` Search Term _(ex. "foo")_
3. `String` Search Column (ex. "A:A")
<br>

**Output**: `Integer` Row Number

## 