Declare variable "myStr" as a string variable type.
```vba
Sub Macro1()

Dim myStr as string

End Sub
```
## Types of Variables in VBA
| Name      | Type      | Explanation                                                                                                                                                                                                                                      |
|-----------|-----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Numerical | Integer   | Accepts only integer values, mainly used for counters; value needs to be between -32768 and 32767. Note: You should always use Long instead of Integer. Integer numbers used to be needed to reduce memory usage. But it is no longer necessary. |
| Numerical | Long      | Accepts only integer values, used for larger referencing like populations; value needs to be between -2,147,483,648 and 2,147,483,648                                                                                                            |
| Numerical | Double    | Accepts decimal values with significant degree of precision; values need to be between -1.79769313486231e308and -4.94065645841247e-324 for negative numbers and 1.79769313486231e308and 4.94065645841247e-324 for positive numbers.              |
| Text      | String    | Accepts strings of text, usually identified with double quotation marks; if a value is input without quotation marks, it will be automatically recognised as text.                                                                               |
| Date/Time | Date      | Accepts dates, needs to be between # signs, e.g. #31/12/1999#                                                                                                                                                                                    |
| Boolean   | Boolean   | Accepts True or False values.                                                                                                                                                                                                                    |
| Any       | Variant   | Accepts any type of variable.                                                                                                                                                                                                                    |
| Objects   | Workbook  | Accepts workbook names.                                                                                                                                                                                                                          |
| Objects   | Worksheet | Accepts worksheet names.                                                                                                                                                                                                                         |
| Objects   | Object    | Accepts all objects                                                                                                                                                                                                                              |

Set cell A1 equal to the variable j.
```vba
Sub Macro1()
  Dim j as long
  j = 2
  j = j + 1

range("A1").value = j

End Sub
```
