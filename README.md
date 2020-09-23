<div align="center">

## Total Record Count In ADO Recordset


</div>

### Description

This simple little function just returns the total number of records in a ADO recordset.
 
### More Info
 
A ADODB.Recordset structure

A Long integer with the total number of records

I believe it's not the fastest way of retrieving the information but at least it works. I'd like it if somebody can suggest a alt. way of getting to this info.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |4.2 (162 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-total-record-count-in-ado-recordset__1-1742/archive/master.zip)





### Source Code

```
'Pass this function your ADO recordset
Function GetTotalRecords(ByRef aRS As ADODB.Recordset) As Long
On Error GoTo handelgettotalrec
 Dim adoBookM As Variant 'Declare a variable to keep the current location
 adoBookM = aRS.Bookmark 'Get the current location in the recordset
 aRS.MoveLast   'Move to the last record in the recordset
 GetTotalRecords = aRS.RecordCount 'Set the count value
 aRS.Bookmark = adoBookM 'Return to the origanal record
 Exit Function
handelgettotalrec:
 GetTotalRecords = 0  'If there's any errors return 0
 Exit Function
End Function
```

