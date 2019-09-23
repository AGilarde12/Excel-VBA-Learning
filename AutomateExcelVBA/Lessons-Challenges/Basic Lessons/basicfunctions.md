# Basic Functions of VBA - 

Setting a range to value:
```vba
Sub Macro1()

range("A2:B3") = 5

End Sub
```

Setting a range's values to a formula:
```vba
Sub Macro1()

range("A2:A3").formula = "=5*2"

End Sub
```

Set cell A2 = B2 
```vba
Sub Macro1()

range("A2").value = range("B2").value

End Sub
```

Referencing other workbooks:
