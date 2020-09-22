<div align="center">

## AS400 to Excel


</div>

### Description

This code uses a value in cell B2 in Excel as a lookup on the AS400 and returns values to cells C2 and D2 in Excel

You must add a reference to:

Microsoft ActiveX DAta Objects 2.0 library
 
### More Info
 
Put the value you would like to lookup in cell B2

then run macro

DSN-Less connection


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Belluscio](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-belluscio.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VBA MS Excel
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-belluscio-as400-to-excel__1-26024/archive/master.zip)





### Source Code

```
Dim CN As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim strSQL As String
Sub GetNameCity2()
CN.Open "Driver={Client Access ODBC Driver (32-bit)}; System=typeyouras400ipaddress-or-as400namehere; Uid=typeyouras400Namehere; Pwd=typeyouras400passwordhere;" ' open connection to database
'this section retrieves the name and site
'PLTFILES# is the library
'ONETI561 is the file
'NAME, CITY, ADRNUM are the fields to retrieve
RS.Open strSQL, CN
strSQL = "select NAME, CITY, ADRNUM from PLTFILES#.ONETI561 where PRADDR = 'Y' AND ADRNUM = '" & Range("B2").Value & "'"
RS.Open strSQL, CN
If RS.BOF Or RS.EOF Then
 msgbox "Could not find lookup value."
Else
 RS.MoveFirst
 Range("C2").Value = RS.Fields(0)
 Range("D2").Value = RS.Fields(1)
End If
RS.Close 'Close recordset
CN.Close 'Close connection
End Sub
```

