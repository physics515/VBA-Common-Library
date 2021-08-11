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

## Damerau-Levenshtein Distance (String Metric)
This function takes two strings of any length and calculates the Damerau-Levenshtein Distance between them. Damerau-Levenshtein Distance differs from Levenshtein Distance in that it includes an additional operation, called Transpositions, which occurs when two adjacent characters are swapped. Thus, Damerau-Levenshtein Distance calculates the number of Insertions, Deletions, Substitutions, and Transpositons needed to convert string1 into string2. As a result, this function is good when it is likely that spelling errors have occured between two string where the error is simply a transposition of 2 adjacent characters.

**Required Types**:
1. `common.CaseSensitivity`

**Required References**:
1. `Microsoft Scripting Library`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_
<br> 
**Output**: `Integer` Distance _(ex. "5")_

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

## Fuzzy Find
Configurable fuzzy find algorithm for string matching.

**Required Types**:
1. `common.CaseSensitivity`

**Required References**:
1. `Microsoft Scripting Library`

**Required Functions**:
1. `common.originalMetric`
2. `common.damerau`
3. `common.hamming`
4. `common.levenshtein`
5. `common.sorensenDice`
6. `common.ngrams`
7. `common.tversky`
8. `common.uniqueArrayElements`
9. `common.jaccard`
10. `common.jaroWinkler`
11. `common.simpleMatching`
12. `common.min`
13. `common.max`

**Input**:
1. `String` Query _(ex. "foo")_
2. `Range` Search Range _(ex. Range("A1:B5"))_
3. `Worksheet` Search Sheet Name _(ex. ThisWorkbook.Sheets("Sheet 1"))_
4. Optional `CaseSensitivity` Case Sensitive _(ex. CaseSensitivity.Sensitive)_
5. Optional `Variant` Weights _(ex. Array(1, .2, 3, 4, 5, .06, 7, 8, .009))_
6. Optional `Boolean` Tversky Symmetry _(ex. True)_
7. Optional `Variant` Tversky Weights _(ex. Array(1, 2))_

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

## Hamming Distance (String Metric)
This function takes two strings of the same length and calculates the Hamming Distance between them. Hamming Distance measures how close two strings are by checking how many Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers.

**Required Types**:
1. `common.CaseSensitivity`

**Required References**:
1. `Microsoft Scripting Library`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_
<br> 
**Output**: `Integer` Distance _(ex. "5")_

<br> 

## HTTP Request
Sends http POST or GET request and returns the response.

**Input**:
1. `String` URL _(ex. "https://api.insightly.com/v3.1/Contacts/")_
2. `Boolean` Post _(ex. `True` or `False`)_  - Optional

**Output**: `String` HTTP Response

<br> 

## Jaccard Similarity Coefficient (String Metric)
Calculate the Jaccard Similarity Coefficient.

**Required Types**:
1. `common.CaseSensitivity`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_

**Output**: `Double` Coefficient _(ex. ".1234")_

<br> 

## Jaro-Winkler Distance (String Metric)

Calculate the Jaro-Winkler distance.

**Required Types**:
1. `common.CaseSensitivity`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_

**Output**: `Double` Distance _(ex. ".1234")_

<br> 

## JSON Converter _(Class)_
Tools for using JSON with VBA.

_VBA-JSON v2.3.1 - (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON_

_JSONLib - Copyright (c) 2013, Ryo Yokoyama http://code.google.com/p/vba-json/_

_VBA-UTC v1.0.6 - (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter_

<br> 

### Convert to ISO
Convert local date to ISO 8601 string.

**Input**:  `Date` UTC Local Date
<br> 
**Output**: `String` ISO Date

<br> 

### Convert to JSON
Convert dictionary, collection, or array to JSON.

**Input**:
1. `Variant` Dictionary, Collection, or Array to be converted
2. `Variant` Whitespace - "Pretty" print json with given number of spaces per indentation (Integer) or given string
3. `Long` Current indentation (Default: `0`)

**Output**: `String` JSON

<br> 

### Convert to UTC
Convert local date to UTC date.

**Input**: `Date` Local Date
<br> 
**Output**: `Date` UTC Date

<br> 

### Parse ISO
Parse ISO 8601 date string to local date.

**Input**: `String` ISO date string
<br> 
**Output**: `Date` Local Date

<br> 

### Parse JSON
Convert JSON to dictionary or collection.

**Input**: `String` JSON
<br> 
**Output**: `Object` Dictionary or Collection

<br> 

### Parse UTC
Parse UTC date to local date.

**Input**: `Date` UTC Date
<br> 
**Output**: `Date` Local Date

<br> 

## Levenshtein Distance (String Metric)
This function takes two strings of any length and calculates the Levenshtein Distance between them. Levenshtein Distance measures how close two strings are by checking how many Insertions, Deletions, or Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers. Unlike Hamming Distance, Levenshtein Distance works for strings of any length and includes 2 more operations. However, calculation time will be slower than Hamming Distance for same length strings, so if you know the two strings are the same length, its preferred to use Hamming Distance.

**Required Types**:
1. `common.CaseSensitivity`

**Required References**:
1. `Microsoft Scripting Library`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_

**Output**: `Integer` Distance _(ex. "5")_

<br> 

## Lock All Sheets _(Sub-Routine)_
Locks or unlocks all sheets, unless a sheet is provided then only that sheet will be locked or unlocked.

**Input**:
1. `Boolean` Locked _(ex. `True` or `False`)_
2. `Worksheet` Single Sheet

**Output**: `None`

<br> 

## Maximum Value In An Array (max)
This function takes multiple numbers or multiple arrays of numbers and returns the max number. This function also accounts for numbers that are formatted as strings by converting them into numbers.

**Input**: `Variant` Numbers _(ex. Array(1, 3, 5, 5, 9, 9.5))_

**Output**: `Double` Max Value _(ex. 9.5)_

<br> 

## Minimum Value In An Array (min)
This function takes multiple numbers or multiple arrays of numbers and returns the min number. This function also accounts for numbers that are formatted as strings by converting them into numbers.

**Input**: `Variant` Numbers _(ex. Array(.5, 1, 3, 5, 5, 9, 9.5))_

**Output**: `Double` Min Value _(ex. .5)_

<br> 

## nGrams
Determine the grams of a given length for a string. (ex. nGrams("Hello World", 2) = ("He", "el", "ll", "lo", "o ", " W", "Wo", "or", "rl", "ld")

**Input**:
1. `String` text _(ex. "foo")_

**Output**: `Variant` nGram _(ex. Array("He", "el", "ll", "lo", "o ", " W", "Wo", "or", "rl", "ld"))_

<br> 

## One Digit Number to Text
Converts any single-digit number to text.

**Input**: `String` Digit _(ex. "5")_
<br> 
**Output**: `String` Text _(ex. "Five")_

<br> 

## Original Metric (String Metric)
String metric.

**Required Types**:
1. `common.CaseSensitivity`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_
<br> 
**Output**: `Double` Metric _(ex. ".5")_

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

## Simple Matching Metric (String Metric)
Calculate the simple matching metric.

**Required Types**:
1. `common.CaseSensitivity`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_

**Output**: `Double` Metric _(ex. ".1234")_

<br> 

## Sorensen-Dice Distance (String Metric)
Get the edit-distance according to Dice between two values.

**Required Types**:
1. `common.CaseSensitivity`

**Required References**:
1. `Microsoft Scripting Library`

**Required Functions**:
1. `common.ngrams`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_
<br> 
**Output**: `Integer` Distance _(ex. "5")_

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

<br>

## Tversky Index (String Metric)
Computes the Tversky index between two sequences. For alpha = beta = 0.5, the index is equal to Dice's coefficient. For alpha = beta = 1, the index is equal to the Tanimoto coefficient.

**Required Types**:
1. `common.CaseSensitivity`

**Required Functions**:
1. `common.uniqueArrayElements`

**Input**:
1. `String` String1 _(ex. "foo")_
2. `String` String2 _(ex. "bar")_
3. `CaseSensitive` Case Sensitivity _(ex. CaseSensitive.Sensitive)_
4. Optional `Boolean` Symmetry _(ex. True)_
5. Optional `Double` String1 Weight _(ex. .5)_
6. Optional `Double` String2 Weight _(ex. .5)_

**Output**: `Double` Index _(ex. ".1234")_

<br> 

## Unique Array Elements
Computes the unique elements of an array.

**Required References**:
1. `Microsoft Scripting Library`

**Input**: `Variant` Array _(ex. Array(1,1,2,3))_

**Output**: `Variant`  Array _(ex. Array(1,2,3))_

<br> 