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

The V&C Timesheeter allows for template-based, automated creation of Excel timesheets. Future timesheets may need to be dupplicated, edited or adjusted for various reasons, and the manual creation and updating of these timesheets can be a labourous task. V&C Timehsheeter saves the user time and effort by automating this process. Simply supply a blank timesheet timeplate with the required updates and V&C Timesheeter will produce a full set of updated timesheets for the given period. 

## 📔 Detailed Usage Instructions: Read them here! → Documentation

## ⚡️ Quick Start

### Basic Usage

1. Prepare a `blank.xlsx` timesheet Excel file. This file should include a `Date` sheet with the timesheet period's `starting date`, where subsequent sheets will rely on this cell to create their date ranges.
2. Make sure the provided or compiled `timesheeter.exe` program and your `blank.xlsx` file are in the same directory.
3. Open `timesheeter.exe`, drag and drop `blank.xlsx` into the specified area.
4. Edit the form fields as necessary (see Instructions for more info).
5. Press <i>Go!</i>. Your created timesheets should now be in the specified output path.

### Compiling `timesheeter.exe` (For Developers)

If you do not have a provided copy of the `timesheeter.exe` program, you can either compile your own or contact the developer and request a new copy.
1. Clone this repo and do `pip install -r requirements.txt`
2. Run `pyinstaller ...`

## 📝 Contact Me

If you need to reach me, feel free at larryhiehuynh@gmail.com


