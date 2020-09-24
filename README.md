<div align="center">

## Database paging in ASP


</div>

### Description

I wanted to add paging code in my project.I've seen all other codes on the site but they are not at all worth for me. Now I've written a code which is very easy to understand and can be used by any student or professional in their projects.
 
### More Info
 
This program assumes that you have dsn = 'myDSN' or replace 'myDSN' with your existing DSN.

This program displays 5 records at a time from any database. You can change no. of records per page by changing the value of iPageSize variable


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vivek Tanaji Kulthe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vivek-tanaji-kulthe.md)
**Level**          |Intermediate
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vivek-tanaji-kulthe-database-paging-in-asp__4-8122/archive/master.zip)





### Source Code

```
<%
Const iPageSize=5	'How many records to show
Dim CPage			'Current Page No.
Dim Cn				'Connection Object
Dim Rs				'Recordset Object
Dim TotPage			'Total No. of pages if iPageSize records are displayed per page.
Dim i				'Counter
CPage=Cint(Request.Form("CurrentPage"))	'get CPage value from form's CurrentPage field
Select Case Request.Form("Submit")
	Case "Previous"						'if prev button pressed
		CPage = Cint(CPage) - 1			'decrease current page
	Case "Next"							'if next button pressed
		CPage = Cint(CPage) + 1			'increase page count
End Select
Set	Cn=Server.CreateObject("ADODB.Connection")	'create connection
	Cn.CursorLocation = 3
	Cn.Open "myDSN"
Set	Rs=Server.CreateObject("ADODB.Recordset")	'create recordset
	Rs.Open "Select * from studentmaster",Cn,2,2
	Rs.PageSize=iPageSize
If CPage=0 then CPage=1						'initially make current page = first page
If Not(Rs.EOF) Then Rs.AbsolutePage=CPage	'specifies that current record resides in CPage
TotPage=Rs.PageCount						'stores total no. of pages
%>
<HTML>
<BODY>
<H2>Database paging example</H2>
by Vivek Kulthe (<a href = "mailto:vivekkulthe@yahoo.com">vivekkulthe@yahoo.com</a>)<P>
<TABLE BORDER = 1>
<%
Response.Write("<TR><TD><B>" & Rs.Fields(1).Name & "</TD><TD><B>" & Rs.Fields(2).Name	& "</TD></TR>")	'display title for table
%>
<%
For i=1 to Rs.PageSize
	Response.Write ("<TR><TD>" & Rs(1) & "</TD><TD>" & Rs(2) & "</TD><TR>")	'display table records upto PageSize
	Rs.MoveNext
	If Rs.EOF Then Exit For
Next
'close all connections and recordsets
Rs.Close
Cn.Close
Set Rs = Nothing
Set Cn = Nothing
%>
</TABLE>
<BR>
Page <%=CPage %> of <%=TotPage %><p>
<!--'store current page value in hidden type and display next-prev buttons-->
<FORM Action="<%=Request.ServerVariables("SCRIPT_NAME") %>" Method=POST>
		<Input Type=Hidden name="CurrentPage" Value="<%=CPage%>" >
	<% If CPage > 1 Then %>
		<Input type=Submit Name="Submit" Value="Previous">
	<% End IF%>
	<% If CPage <> TotPage Then %>
		<Input type=Submit Name="Submit" Value="Next">
	<% End If %>
</FORM>
</BODY>
</HTML>
```

