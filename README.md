# VBScript Get-ChildItem: Missing Recursion
This example demonstrates a common error in VBScript when attempting to recursively traverse directories using Get-ChildItem.  Unlike PowerShell, VBScript's Get-ChildItem does not support a direct equivalent to the "-Recurse" parameter.

## Bug
The bug lies in the naive attempt to use a parameter that doesn't exist.  This will cause a runtime error.

## Solution
The solution involves a recursive function to manually traverse the directory structure.
