<div align="center">

## Get the ID of just inserted database record using INSERT clause\.


</div>

### Description

This article demonstrates how to get the ID of just inserted database record using INSERT clause.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sergei Belov](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sergei-belov.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sergei-belov-get-the-id-of-just-inserted-database-record-using-insert-clause__4-7379/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<HTML>
<HEAD>
<TITLE>Document Title</TITLE>
</HEAD>
<BODY>
<%
' Please note that this example won't work with MS Access database since @@IDENTITY is only suported by SQL Server.
'
Dim loCN ' Connection Object
Dim loRS ' Recordset Object
Dim lsSQL ' SQL String
Dim lnNewEmployeeID ' New ID
'
'# Create a database connection
Set loCN = Server.CreateObject("ADODB.Connection")
Call loCN.Open("DSN=Northwind", "sa", "password")
'
'# Build INSERT statement
lsSQL = "INSERT INTO Employees " & _
 "(FirstName, LastName, Title, HomePhone) " & _
 "VALUES ('John','Doe','Some Title','123-456-7890')"
'
'# Append SELECT statement with identity which we will use later to retrieve the idetity of the new record
lsSQL = lsSQL & " SELECT * FROM Employees WHERE " & _
 "EmployeeID = @@IDENTITY"
'
'# Execute SQL
Set loRS = loCN.Execute(lsSQL).NextRecordset
'
'# Here is an alternative to the line
' above
' Set loRS = Server.CreateObject("ADODB.Recordset")
' Call loRS.Open(lsSQL, loCN)
' Set loRS = loRS.NextRecordset
'
'# Get the ID of the record we
' just inserted
lnNewEmployeeID = loRS("EmployeeID").Value
'
Response.Write lnNewEmployeeID
'
' Clean Up
Set loRS = Nothing
Set loCN = Nothing
'
%>
</BODY>
</HTML>
```

