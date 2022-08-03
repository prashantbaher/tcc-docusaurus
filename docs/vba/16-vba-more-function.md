---
title: VBA Functions that do more
tags:   [VBA]
permalink: /vba/more-functions/
---

A few VBA `functions` go above and beyond the call of duty. Rather than simply return a value, these functions have some useful side effects. 

Below table lists them.

<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">Functions with Useful Side Benefits</th>
    </tr>
    <tr>
        <th>Function</th>
        <th>What is does</th>
    </tr>
    <tr>
        <td>MsgBox</td>
        <td>
            Displays a handy dialog box containing a message and buttons. 
            The function returns a code that identifies which button the user clicks.
        </td>
    </tr>
    <tr>
        <td>InputBox</td>
        <td>
            Displays a simple dialog box that asks the user for some input. 
            The function returns whatever the user enters into the dialog box.
        </td>
    </tr>
    <tr>
        <td>Shell</td>
        <td>
            Executes another program. The function returns the task ID (a unique identifier) 
            of the other program (or an error if the function can’t start the other program).
        </td>
    </tr>
    <tr>
        <td>GetObject/CreateObject</td>
        <td>
            Returns/Create a reference to an object provided by an ActiveX component. 
            (If you don't understand, don't bother about it. Just remember we use this 
            function to for checking &#38; creating objects in later topics)
        </td>
    </tr>
</table>

---

## Discovering VBA functions

How do we find out which function does VBA provides? 

The best source is the *Visual Basic Help system* in build in your CAD Application. 

I compiled a partial list of `functions`, which I share with you in following Table. 

I omitted some of the more specialized or obscure functions. 

For complete details on a particular function, type the function name into a VBA module, move the cursor anywhere in the text, and press `F1`. 

<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">VBA’s Most Useful Built-In Functions</th>
    </tr>
    <tr>
        <th>Function</th>
        <th>What is does</th>
    </tr>
    <tr>
        <td>Abs</td>
        <td>Returns a number’s absolute value.</td>
    </tr>
    <tr>
        <td>Array</td>
        <td>Returns a variant containing an array.</td>
    </tr>
    <tr>
        <td>Asc</td>
        <td>Converts the first character of a string to its ASCII value.</td>
    </tr>
    <tr>
        <td>Atn</td>
        <td>Returns the arctangent of a number.</td>
    </tr>
    <tr>
        <td>Choose</td>
        <td>Returns a value from a list of items.</td>
    </tr>
    <tr>
        <td>Chr</td>
        <td>Converts an ANSI value to a string.</td>
    </tr>
    <tr>
        <td>Cos</td>
        <td>Returns a number’s cosine.</td>
    </tr>
    <tr>
        <td>CurDir</td>
        <td>Returns the current path.</td>
    </tr>
    <tr>
        <td>Date</td>
        <td>Returns the current system date.</td>
    </tr>
    <tr>
        <td>DateAdd</td>
        <td>Returns a date to which a specified time interval has been added — for example, one month from a particular date.</td>
    </tr>
    <tr>
        <td>DatePart</td>
        <td>Returns an integer containing the specified part of a given date — for example, a date’s day of the year.</td>
    </tr>
    <tr>
        <td>DateSerial</td>
        <td>Converts a date to a serial number.</td>
    </tr>
    <tr>
        <td>DateValue</td>
        <td>Converts a string to a date.</td>
    </tr>
    <tr>
        <td>Day</td>
        <td>Returns the day of the month from a date value.</td>
    </tr>
    <tr>
        <td>Dir</td>
        <td>Returns the name of a file or directory that matches a pattern.</td>
    </tr>
    <tr>
        <td>Erl</td>
        <td>Returns the line number that caused an error.</td>
    </tr>
    <tr>
        <td>Err</td>
        <td>Returns the error number of an error condition.</td>
    </tr>
    <tr>
        <td>Error</td>
        <td>Returns the error message that corresponds to an error number.</td>
    </tr>
    <tr>
        <td>Exp</td>
        <td>Returns the base of the natural logarithm (e) raised to a power.</td>
    </tr>
    <tr>
        <td>FileLen</td>
        <td>Returns the number of bytes in a file.</td>
    </tr>
    <tr>
        <td>Fix</td>
        <td>Returns a number’s integer portion.</td>
    </tr>
    <tr>
        <td>Format</td>
        <td>Displays an expression in a particular format.</td>
    </tr>
    <tr>
        <td>GetSetting</td>
        <td>Returns a value from the Windows registry.</td>
    </tr>
    <tr>
        <td>Hex</td>
        <td>Converts from decimal to hexadecimal.</td>
    </tr>
    <tr>
        <td>Hour</td>
        <td>Returns the hours portion of a time.</td>
    </tr>
    <tr>
        <td>InputBox</td>
        <td>Displays a box to prompt a user for input.</td>
    </tr>
    <tr>
        <td>InStr</td>
        <td>Returns the position of a string within another string.</td>
    </tr>
    <tr>
        <td>Int</td>
        <td>Returns the integer portion of a number.</td>
    </tr>
    <tr>
        <td>IPmt</td>
        <td>Returns the interest payment for an annuity or loan.</td>
    </tr>
    <tr>
        <td>IsArray</td>
        <td>Returns True if a variable is an array.</td>
    </tr>
    <tr>
        <td>IsDate</td>
        <td>Returns True if an expression is a date.</td>
    </tr>
    <tr>
        <td>IsEmpty</td>
        <td>Returns True if a variable has not been initialized.</td>
    </tr>
    <tr>
        <td>IsError</td>
        <td>Returns True if an expression is an error value.</td>
    </tr>
    <tr>
        <td>IsMissing</td>
        <td>Returns True if an optional argument was not passed to a procedure.</td>
    </tr>
    <tr>
        <td>IsNull</td>
        <td>Returns True if an expression contains no valid data.</td>
    </tr>
    <tr>
        <td>IsNumeric</td>
        <td>Returns True if an expression can be evaluated as a number.</td>
    </tr>
    <tr>
        <td>IsObject</td>
        <td>Returns True if an expression references an OLE Automation object.</td>
    </tr>
    <tr>
        <td>LBound</td>
        <td>Returns the smallest subscript for a dimension of an array.</td>
    </tr>
    <tr>
        <td>LCase</td>
        <td>Returns a string converted to lowercase.</td>
    </tr>
    <tr>
        <td>Left</td>
        <td>Returns a specified number of characters from the left of a string.</td>
    </tr>
    <tr>
        <td>Len</td>
        <td>Returns the number of characters in a string.</td>
    </tr>
    <tr>
        <td>Log</td>
        <td>Returns the natural logarithm of a number to base.</td>
    </tr>
    <tr>
        <td>LTrim</td>
        <td>Returns a copy of a string, with any leading spaces removed.</td>
    </tr>
    <tr>
        <td>Mid</td>
        <td>Returns a specified number of characters from a string.</td>
    </tr>
    <tr>
        <td>Minutes</td>
        <td>Returns the minutes portion of a time value.</td>
    </tr>
    <tr>
        <td>Month</td>
        <td>Returns the month from a date value.</td>
    </tr>
    <tr>
        <td>MsgBox</td>
        <td>Displays a message box and (optionally) returns a value.</td>
    </tr>
    <tr>
        <td>Now</td>
        <td>Returns the current system date and time.</td>
    </tr>
    <tr>
        <td>RGB</td>
        <td>Returns a numeric RGB value representing a color.</td>
    </tr>
    <tr>
        <td>Replace</td>
        <td>Replaces a substring in a string with another substring.</td>
    </tr>
    <tr>
        <td>Right</td>
        <td>Returns a specified number of characters from the right of a string.</td>
    </tr>
    <tr>
        <td>Rnd</td>
        <td>Returns a random number between 0 and 1.</td>
    </tr>
    <tr>
        <td>RTrim</td>
        <td>Returns a copy of a string, with any trailing spaces removed.</td>
    </tr>
    <tr>
        <td>Second</td>
        <td>Returns the seconds portion of a time value.</td>
    </tr>
    <tr>
        <td>Sgn</td>
        <td>Returns an integer that indicates a number’s sign.</td>
    </tr>
    <tr>
        <td>Shell</td>
        <td>Runs an executable program.</td>
    </tr>
    <tr>
        <td>Sin</td>
        <td>Returns a number’s sine.</td>
    </tr>
    <tr>
        <td>Space</td>
        <td>Returns a string with a specified number of spaces.</td>
    </tr>
    <tr>
        <td>Split</td>
        <td>Splits a string into parts, using a delimiting character.</td>
    </tr>
    <tr>
        <td>Sqr</td>
        <td>Returns a number’s square root.</td>
    </tr>
    <tr>
        <td>Str</td>
        <td>Returns a string representation of a number.</td>
    </tr>
    <tr>
        <td>StrComp</td>
        <td>Returns a value indicating the result of a string comparison.</td>
    </tr>
    <tr>
        <td>String</td>
        <td>Returns a repeating character or string.</td>
    </tr>
    <tr>
        <td>Tan</td>
        <td>Returns a number’s tangent.</td>
    </tr>
    <tr>
        <td>Time</td>
        <td>Returns the current system time.</td>
    </tr>
    <tr>
        <td>Timer</td>
        <td>Returns the number of seconds since midnight.</td>
    </tr>
    <tr>
        <td>TimeSerial</td>
        <td>Returns the time for a specified hour, minute, and second.</td>
    </tr>
    <tr>
        <td>TimeValue</td>
        <td>Converts a string to a time serial number.</td>
    </tr>
    <tr>
        <td>Trim</td>
        <td>Returns a string without leading or trailing spaces.</td>
    </tr>
    <tr>
        <td>TypeName</td>
        <td>Returns a string that describes a variable’s data type.</td>
    </tr>
    <tr>
        <td>UBound</td>
        <td>Returns the largest available subscript for an array’s dimension.</td>
    </tr>
    <tr>
        <td>UCase</td>
        <td>Converts a string to uppercase.</td>
    </tr>
    <tr>
        <td>Val</td>
        <td>Returns the numbers contained in a string.</td>
    </tr>
    <tr>
        <td>VarType</td>
        <td>Returns a value indicating a variable’s subtype.</td>
    </tr>
    <tr>
        <td>Weekday</td>
        <td>Returns a number representing a day of the week.</td>
    </tr>
    <tr>
        <td>Year</td>
        <td>Returns the year from a date value.</td>
    </tr>
</table>

