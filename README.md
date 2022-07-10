<p align="center">
  <h1 align="center" style="margin: 0 auto 0 auto;">V&C Timesheeter</h1>
  <h3 align="center" style="margin: 0 auto 0 auto;">Automated Excel Timesheet Creation</h3>
  <h5 align="center" style="margin: 0 auto 0 auto;">Developed for Vim & Co.</h5>
</p>

<p align="center">
  <img src="https://img.shields.io/github/last-commit/LarryHH/VC_Timesheeter">
  <img src="https://img.shields.io/github/contributors/LarryHH/VC_Timesheeter">
  <img src="https://img.shields.io/github/issues/LarryHH/VC_Timesheeter">
  <img src="https://img.shields.io/github/stars/LarryHH/VC_Timesheeter">
</p>

___

## 💡 Introduction

The V&C Timesheeter allows for template-based, automated creation of Excel timesheets. The manual creation or duplication of timesheets for the correct payroll period can be a labourous task. Further, future timesheets may need to be duplicated, edited or adjusted for various reasons, which compounds this labour by many magnitudes. The V&C Timehsheeter saves the user time and effort by automating this process:
- Now, timesheets are created with the correct payroll dates for the whole year all at once.
- Any updates to timesheets (e.g. roster changes) only need to be applied once, rather than for every subsequent timesheet.

Simply supply a blank timesheet timeplate with the required updates and V&C Timesheeter will produce a full set of updated timesheets for the given period. 

<p align="center">
  <img src="/docs/assets/basic_usage.gif" width="80%" height="80%" alt="animated" />
</p>

## 📔 Detailed Usage Instructions: Read them here! → Documentation

## ⚡️ Quick Start

### Basic Usage

1. Prepare a `blank.xlsx` timesheet Excel file. This file should include a `Date` sheet with the timesheet period's `starting date`, where subsequent sheets will rely on this cell to create their date ranges.
2. Make sure the provided or compiled `timesheeter.exe` program and your `blank.xlsx` file are in the same directory.
3. Open `timesheeter.exe`, drag and drop `blank.xlsx` into the specified area.
4. Edit the form fields as necessary (see Instructions for more info).
5. Press <i>Go!</i>. Your created timesheets should now be in the specified output path.

### Compiling `timesheeter.exe` (For Developers)

If you do not have a provided copy of the `timesheeter.exe` program, you can either compile your own or [contact the developer](#contact) and request a new copy.
1. Clone this repo and do `pip install -r requirements.txt`
2. Run `pyinstaller ...`

## ✉️ Contact Me
<a id="contact"></a>

If you need to reach me, feel free at larryhiehuynh@gmail.com


