---
title: VBA Assignment Statements And Operators
tags:   [VBA]
permalink: /vba/assignment-statement-and-operator/
---

An *assignment statement* is a *VBA statement* that assigns the result of an expression to a variable or an object. 

In a book I read Excel’s Help system defines the term expression as: 

> "Combination of keywords, operators, variables, and constants that yields a string, number, or object. An expression can be used to perform a calculation, manipulate characters, or test data." 

Much of your work in VBA involves *developing (and debugging)* expressions. 

If you know how to create simple formulas in Excel, you’ll have no trouble creating expressions. 

With a formula, Excel displays the result in a cell. 

A VBA expression, on the other hand, can be assigned to a variable. 

> For understanding purpose, I used Excel as an example. Please don't get confused with it. 

In the assignment statement examples that follow, the expressions are to the right of the equal sign: 

```vb
X = 1
X = x + 1
X = (y * 2) / (z * 2)
NumberOfParts = 15
SelectObject = True
```

**Expressions** can be as complex as you need them to be; use the line continuation character (a space followed by an underscore) to make lengthy expressions easier to read. 

# Operators

As you can see in the VBA uses the equal sign as its *assignment operator*. 

You’re probably accustomed to using an `equal` sign as a mathematical symbol for equality. 

Therefore, an assignment statement like the following may cause you to raise your eyebrows: 

```vb
x = x + 1
```

How can the variable `x` be equal to itself plus 1? 

Answer: It can’t. 

In this case, the assignment statement is increasing the value of `x` by **1**. 

Just remember that an assignment uses the *equal* sign as an `operator`, not a symbol of equality. 

{% include adsense/adsense-horizontal-ads.html %}

## Smooth Operators

`Operators` play a major role in VBA. Besides the assignment operator i.e. equal sign (discussed in the previous topic), VBA provides several other operators. 

Below table lists these operators. 

<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">VBA’s Operators</th>
    </tr>
    <tr>
        <th>Function</th>
        <th>Operator Symbol</th>
    </tr>
    <tr>
        <td>Addition</td>
        <td>+</td>
    </tr>
    <tr>
        <td>Multiplication</td>
        <td>*</td>
    </tr>
    <tr>
        <td>Division</td>
        <td>/</td>
    </tr>
    <tr>
        <td>Subtraction</td>
        <td>-</td>
    </tr>
    <tr>
        <td>Exponentiation</td>
        <td>^</td>
    </tr>
    <tr>
        <td>String concatenation</td>
        <td>&#38;</td>
    </tr>
    <tr>
        <td>Integer division (the result is always an integer)</td>
        <td>\</td>
    </tr>
    <tr>
        <td>Modulo arithmetic (returns the remainder of a division operation)</td>
        <td>Mod</td>
    </tr>
</table>

The term **concatenation** is programmer speak for “join together”. 

Thus, if you concatenate strings, you are combining strings to make a new and improved string. 

VBA also provides a full set of logical operators. Below table, shows some of logical operators. 

<table class="w3-table-all w3-mobile w3-card-4">
    <tr>
        <th class="w3-center" colspan="2">VBA’s Logical Operators</th>
    </tr>
    <tr>
        <th>Operator</th>
        <th>What is does</th>
    </tr>
    <tr>
        <td>Not</td>
        <td>Performs a logical negation on an expression.</td>
    </tr>
    <tr>
        <td>And</td>
        <td>Performs a logical conjunction on two expressions.</td>
    </tr>
    <tr>
        <td>Or</td>
        <td>Performs a logical disjunction on two expressions.</td>
    </tr>
    <tr>
        <td>XoR</td>
        <td>Performs a logical exclusion on two expressions.</td>
    </tr>
    <tr>
        <td>Eqv</td>
        <td>Performs a logical equivalence on two expressions.</td>
    </tr>
    <tr>
        <td>Imp</td>
        <td>Performs a logical implication on two expressions.</td>
    </tr>
</table> 

The precedence order for *operators* in VBA is exactly the same as in *Excel formulas*. 

*Exponentiation* has the highest precedence. *multiplication* and *division* come next, followed by *addition* and *subtraction*. 

You can use *parentheses* to change the natural precedence order, making whatever’s operation in parentheses come before any operator. 

Take a look at this code: 

```vb
z = x + 5 * y
```

When this code is executed, what’s the value of `z`? 

If you answered **13**, you get a gold star that proves you understand the concept of operator precedence. 

If you answered **16**, read this: The *multiplication* operation (5 * y) is performed first, and that result is added to `x`. 

If you answered something other than **13** or **16**, I have no comment.

By the way, I can never remember how operator precedence works, so I tend to use parentheses even when they aren’t required. 

For example, in real life I would write that last assignment statement like this: 

```vb
z = x + (5 * y)
```

> Don’t be shy about using *parentheses* even if they aren’t required — especially if doing so makes your code easier to understand. VBA doesn’t care if you use *extra parentheses*. 

Next post will be about ***VBA Arrays***.

