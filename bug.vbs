Late Binding: VBScript's default late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic with COM objects where type checking happens at runtime.

Example:
```vbscript
Set objExcel = CreateObject("Excel.Application")
'No error checking to see if Excel is installed
 objExcel.Visible = True
```

This will throw an error if Excel isn't installed or accessible.

Implicit Type Conversion: VBScript's loose typing can cause unexpected behavior when different data types are combined. For instance, comparing a string to a number might not yield the expected Boolean result.

Example:
```vbscript
if "10" = 10 then
  MsgBox "Equal"
else
  MsgBox "Not Equal"
end if
```

Here, "10" and 10 are implicitly converted but the result is unexpected to some.

Unhandled Exceptions: VBScript lacks robust exception handling mechanisms compared to modern languages.  Errors can halt script execution unexpectedly without any informative message, making debugging harder.

Example:
```vbscript
On Error Resume Next
Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile("nonexistent.txt", 1)
If Err.Number <> 0 Then
    MsgBox "Error opening file: " & Err.Number & " - " & Err.Description
End If
```

This uses error handling but many don't, leading to crashes.