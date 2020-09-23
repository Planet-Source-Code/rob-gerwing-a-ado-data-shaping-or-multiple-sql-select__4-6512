<div align="center">

## A ADO Data Shaping or Multiple SQL Select


</div>

### Description

Do you have slow running mulitple SELECT Statements or long reports to fill on a web page. Use the Microsoft Shape Command. Learn to use ADO 2.1 and greater advance features. This code is great for three things, (1) Very fast way to do multiple SQL select statements and reports. (2) Great for databases not Normalized. (3) Avoids multiple nested single threaded ADO Record Sets loops which are very slow.
 
### More Info
 
Basic ADO Recordset use in Microsoft Active server Pages. Basic Knowledge of SQL "SELECT" statements.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rob Gerwing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob-gerwing.md)
**Level**          |Intermediate
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rob-gerwing-a-ado-data-shaping-or-multiple-sql-select__4-6512/archive/master.zip)





### Source Code

```
<%Response.Buffer = true
Const adOpenForwardOnly = 2
Dim connShape,strShape
Dim objConn,objRS,objStartDate,objEndDate
'ADO CONNECTION
Set connShape = Server.CreateObject("ADODB.Connection")
connShape.Provider = "MSDataShape" 'Tell ADO to expect MSShape Command in SQL Syntax
connShape.Open "DSN=NAME-OF-ODBC-SOURCE" 'Insert your Data Source Name String
'ADO RECORDSET
Set objRS = Server.CreateObject("ADODB.Recordset")
'Shape SQL Syntax
[Available at Microsoft KB Article Q189657]
'WHY DID I USE SHAPE AND NOT A JOIN?
'Look at the 2nd and 3rd SELECTS, I needed to retrieve a record associated to
'order_id BUT the same field name "event_value" and different event_types.
'Working Example -->
	strShape = "SHAPE {Select order_id,f_name,l_name FROM Order_Table"&_
	  " WHERE l_name = 'SMITH' ORDER BY l_name} AS OrderData "&_
		 "APPEND "&_
		 "({ SELECT order_id, event_value, event_type FROM event" &_
		 " WHERE event_type = 'UserStartDate' } " &_
		 " RELATE order_id TO order_id) AS STARTDATE, "&_
		 " ({ SELECT order_id, event_value, event_type FROM event" &_
		 " WHERE event_type = 'UserEndDate' } " &_
		 " RELATE order_id TO order_id) AS ENDDATE"
'Open RecordSet
objRS.OPEN strShape,ConnShape,adOpenForwardOnly
Response.Write "<TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0>"
Do While Not objRS.EOF 'Looping through Parent Record Set
'Take from Parent Record Set
Response.Write("<TR><TD>" & objRS("order_id") & "</TD>")
Response.Write("<TD>" & objRS("l_name") & ",&nbsp;" & objRS("f_name") & "</TD>")
'StartDate 1st Child RecordSet No Loop, EXPECTING ONLY ONE RECORD VALUE
	'Must use "STARTDATE" as reference in SQL "AS STARTDATE"
	Set objStartDate = objRS("STARTDATE").Value
	If objStartDate.Eof = True Then
		Response.Write("<TD>&nbsp;</TD>")
	Else
		Response.Write("<TD>" & objStartDate("event_value") & "</TD>")
	End If
	objStartDate.Close
'EndDate 2nd Child RecordSet Loop Used,EXPECTING MORE MULTIPLE RECORD VALUES
	'Must use "ENDDATE" as reference in SQL "AS ENDDATE"
	Set objEndDate = objRS("ENDDATE").Value
	If objEndDate.Eof = True Then
		Response.Write("<TD>&nbsp;</TD>")
	Else
		Response.Write("<TD>")
			while not objEndDate.Eof
				 Response.write (objEndDate("event_value") & ",")
			objEndDate.MoveNext
			wend
		Response.Write("</TD>")
	End If
		objEndDate.Close
Response.Write "</TR>"
objRS.MoveNext
LOOP
Response.Write "</TABLE>"
'Clean Up
connShape.Close
objRS.Close
Set connShape = Nothing
Set objRS = Nothing
%>
```

