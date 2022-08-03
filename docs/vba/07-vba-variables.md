---
title: VBA Variables
tags:   [VBA]
permalink: /vba/variables/
---

`VBA’s` main purpose is to manipulate data. `VBA` stores the *data* in your computer’s *memory*; it may or may not end up on disk. 

Some *data*, such as *sketch*, resides in `objects`. 

Other *data* is stored in `variables` that you create.

A `variable` is simply a *named storage location* in your computer’s memory. 

You have lots of flexibility in naming your `variables`, so make the `variable` names as descriptive as possible.

You assign a value to a `variable` by using the equal **sign operator**.

The `variable` names in these examples appear on both the left and right sides of the equal signs. 

Note that the last example uses two `variables`.

```vb
x = 1
InterestRate = 0.075
LoanPayoffAmount = 243089
DataEntered = False
x = x + 1
UserName = "Bill Gates"
DateStarted = #3/14/2010#
MyNum = YourNum * 1.25
```

`VBA` enforces a few rules regarding `variable` names:

* You can use *letters, numbers, and some punctuation characters*, but the **first character** must be a letter.
* You **cannot** use any *spaces or periods* in a `variable` name.
* `VBA` does not distinguish between *uppercase* and *lowercase* letters.
* You **cannot** use the following characters in a variable name: **#, $, %, &, or !.**
* `Variable` names can be no longer than *255* characters. Of course, you’re only asking for trouble if you use variable names *255* characters long.

To make `variable` names more *readable*, programmers often use mixed case (for example, *PartDimension*) or the underscore character (part_dimension).

`VBA` has many *reserved* words that you **can’t** use for `variable` names or `procedure` names. 

These include words such as `Sub, Dim, With, End, Next, and For`. 

If you attempt to use one of these words as a `variable`, you may get a compile error (which means your code won’t run. 

So, if an assignment statement produces an *error message*, double-check and make sure that the `variable` name isn’t a *reserved* word.

`VBA` does allow you to create `variables` with names that match names in your `CAD's object model`, such as sketch and part. 

But, obviously, using `names` like that just increases the possibility of getting confused. 

So resist the urge to use a variable named *sketch*, and use something like *swSketch*, *mySketch* or any meaning full name instead.

