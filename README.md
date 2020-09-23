<div align="center">

## Compact Database using JRO \(Jet & Replication objects\)


</div>

### Description

I recently developed a database application and wanted to use only ADO and no DAO. I soon found out that compacting the Jet database was impossible using ADO (until 2.1+ came along that is). This code requires a reference to Microsoft Jet and Replication objects 2.1+ Library (which comes with ADO 2.1+). You can download this update from http://www.microsoft.com/data.
 
### More Info
 
I use this routine in the form_unload sub to compact the current database. If you were to try to compact while there was still an active connection, Jet locking would take over and return an error.

Set the current connection to nothing before compacting (set mcn = nothing).

True or False depending on success of operation

Not aware of any


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Justin Spencer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/justin-spencer.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/justin-spencer-compact-database-using-jro-jet-replication-objects__1-8274/archive/master.zip)

### API Declarations

```
'## Requires reference to Microsoft Jet and Replication objects 2.1+ Library (Standard ADO 2.1+ feature).
public const PASSWORD = "password" 'replace with database password
```


### Source Code

```
'## To use:
private sub command1_click()
  msgbox compressdatabase ("C:\database.mdb") '## Replace with path to database
end sub
Public Function CompressDatabase(mSourceDB As String) As Boolean
on error goto Err
  Dim JRO As JRO.JetEngine
  Set JRO = New JRO.JetEngine
  Dim srcDB As String
  Dim destDB As String
  srcDB = mSource
  destDB = "backup.mdb"
  JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & srcDB & ";Jet OLEDB:Database Password=" & PASSWORD, _
  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & destDB & ";Jet OLEDB:Database Password=" & PASSWORD & ";Jet OLEDB:Engine Type=4"
  Kill srcDB
  DoEvents
  Name destDB As srcDB
  compressdatabase = true
  exit function
Err:
  compressdatabase = false
End Function
```

