# How to use V&C Timesheeter

### [>> Homepage](https://github.com/LarryHH/VC_Timesheeter/blob/master/README.md)

## Table of Contents

1. [**Prerequisites**](prereqs)

2. [**Template Excel File**](excel)
   1. [Accepted File Types](#excel-1)
   2. [Template File Contents](#excel-2)
   3. [Updating the Template File](#excel-3)

3. [**UI - Drag and Drop Window**](#dnd)
   1. [Inputting Your File](#dnd-1)

4. [**UI - Form Window**](#form)
   1. [Form Fields](#form-1)

5. [**Troubleshooting**](#Troubleshooting)


# 1. Prerequisites <small>[[↑](#table-of-contents)]</small>

- Windows 10 or later
- Installed Custom Font: [Engravers Gothic BT Regular](https://www.dafontfree.net/engraversgothic-bt-regular/f123235.htm)

# 2. Template Excel File <small>[[↑](#table-of-contents)]</small>

## 2.1. Accepted File Types

V&C Timesheeter will only accept MS Excel files and associated file types: `xls` and `xlsx`. Other commonly used Excel file types, such as `csv` and `tsv` are not supported.

## 2.2. Template File Contents

The provided `template.xlsx` file must adhere to the following:
- There should exist a sheet named _Date_ (or _date_): the `Date Sheet`.
- The `Date Sheet` sheet should contain only 1 date-typed cell: the `Date Cell`.

The V&C Timesheeter will read the given template file and locate the `Date Cell` within `Date Sheet`.
