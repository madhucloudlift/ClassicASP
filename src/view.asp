<html>
<head>
<title> [ AspBB ] - View </title>
<style type="text/css">
A:link, A:visited { text-decoration: none; color: 'silver' }
</style>
</head>

<body bgcolor="#006699">
<a name='#top'></a>
<%
dim seepage
if request.querystring("seepage") = "" OR request.querystring("seepage") < 1 then
seepage = 1
else
seepage = request.querystring("seepage")
end if
 
id = Request.QueryString("msgID")
set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 

Set rsadd = Server.CreateObject("ADODB.Recordset")
rsadd.Open "SELECT * FROM board WHERE msgID="&id, conn, 1,2
rsadd("threadID") = rsadd("msgID")
rsadd.Update
%>
<div align=center>
<table border=0 width=80%>
<tr>
      <td width="67%"> 
        <%
if Session("currentpage")="" then
Response.Write "<font face=arial size=1><a href='default.asp'>Back to Topics</a>"
else
Response.Write "<font face=arial size=1><a href='default.asp?currentpage="& Session("currentpage") &"'>Back to Topics</a></font>"
end if
%>
      </td>
      <td align=right width="33%"> 
        <%
Set rsp = Server.CreateObject("ADODB.Recordset")
rsp.Open "SELECT * FROM board where msgID>"&id&" AND threadP=0 ORDER BY msgID ASC", conn, 3
if rsp.EOF then
Response.Write "<font face=arial size=1 color=white> << View Pervious Thread :</font>"
else
Response.Write "<font face=arial size=1 color=white><a href='view.asp?msgid="& rsp("msgID") &"'> << View Pervious Thread</a> :</font>"
end if

Set rsn = Server.CreateObject("ADODB.Recordset")
rsn.Open "SELECT * FROM board where msgID<"&id&" AND threadP=0 ORDER BY msgID DESC", conn, 3
if rsn.EOF then
Response.Write "<font face=arial size=1 color=white>: View Next Thread >> </font>"
else
Response.Write "<font face=arial size=1 color=white>: <a href='view.asp?msgid="& rsn("msgID") &"'>View Next Thread >> </a></font>"
end if
%>
      </td>
</tr>
</table>
<br>
<%
Set rsview = Server.CreateObject("ADODB.Recordset")
rsview.CursorLocation = 3 'set cursorlocation to aduserclient 
rsview.CacheSize = 10 'cache the number of record to display per page into cache
rsview.Open "SELECT * FROM board WHERE ThreadID="&id&" ORDER BY msgID", conn
rsview.PageSize = 10 'Set that each page will list 10 records

Dim TotalPages, TotalRecords
TotalPages = rsview.PageCount 'Pagecount will count how many page will need if 10 record fill one page
TotalRecords = rsview.RecordCount
rsview.AbsolutePage = seepage 'the currentpage value will be the rs.Absolute value
Dim count 'Counting variable for our recordset
count = 0 

Do While Not rsview.EOF AND count < rsview.PageSize
%> 
<div align=center>
<table border=1 width=80%cellpadding=0 cellspacing=0>
<tr><td>
<font face=arial size=2 color=white><b><%=rsview("topic")%></b></font>
</td></tr>
</table> 
<table border=0 width=80% cellpadding=0 cellspacing=0>
<tr>
      <td valign=top width="25%">
	  <%  
	  if instr(rsview("email") & "" , "@") > 0 then
	  Response.Write "<font face=arial color=silver size=2><b><a href='mailto:" & trim(rsview("email")) & "'>" & trim(rsview("name")) & "</a></b></font>"
	  else
	  Response.Write "<b><font face=arial color=white size=2>"& rsview("name") &"</b></font>"
	  end if    
	  %>
	  </td>
      <td width="75%"> 
	 <font face=arial size=1 color=silver>Posted on <%=rsview("date")%></font><br>
	 <hr noshade>
	 <font face=arial color=white size=2><b><%=rsview("body")%></b></font><br><br>
     </td>
</tr>
</table>
<table border=0 width=80% cellpadding=0 cellspacing=0>
<tr>
      <td width="25%"><font face=arial size=1 color><b><a href='#top'>Back to Top</a></b></font></td>
      <td width="15%">
	  <%    
	  if instr(rsview("email") & "" , "@") > 0 then
	  Response.Write "<font face=arial color=silver size=2><b><a href='mailto:" & trim(rsview("email")) & "'>E-Mail</a></b></font>"
	  end if  
	  %>
	  </td>
	  <td width="15%">
	  <%    
	  if instr(rsview("http") & "", "http://") > 0 then
      Response.Write "<font face=arial color=silver size=2><b><a href='" & trim(rsview("http")) &"'>WebSite</a></b></font>"
	  end if
      %>
	  </td>
	  <td width="45%"></td>
</tr>
</table>
</div>
<br>
<%
count = count + 1
rsview.MoveNext
Loop

Set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = conn
rs.Open "SELECT * FROM board where msgID="&id, conn, 1, 2
rs("count") = rs("count").OriginalValue + 1
rs.Update
%>
<table border="0" width="84%">
<tr><td align=right>
<font face="arial" size="2">
<a href="repost.asp?msgID=<%=id%>"><b>Reply to this Topic</b></a>
</font>
</td></tr></table>
<br>
<br>
<div align="center">
<table border="0" width="84%">
<tr>
      <td width="21%"><p align="left"><font face="arial" size="2"><b>
	  <%
	 
	  if seepage > 1 then
	  Response.Write "<a href='view.asp?msgID="&id&"&seepage=1'>First</a>"
	  end if
	  
	  %>
	  </b></font></td>
      <td width="21%"><p align="center"><font face="arial" size="2"><b>
	 <% 
	 
	  if seepage > 1 then
	  Response.Write "<a href='view.asp?msgID="&id&"&seepage=" & seepage - 1 &"'><< Pervious</a>"
	  
	  end if
	 %>
	  </b></font></td>
      <td width="21%"><p align="center"><font face="arial" size="2"><b>
	  <% 
	 
	  if CInt(seepage) <> CInt(TotalPages) then
	  Response.Write "<a href='view.asp?msgID="&id&"&seepage=" & seepage + 1 &"'>Next >></a>"
	  
	  end if
	  %>
	  </b></font></td>
      <td width="21%"><p align="right"><font face="arial" size="2"><b>
	  <%
	  
	  if CInt(seepage) <> CInt(TotalPages) then
	  Response.Write "<a href='view.asp?msgID="&id&"&seepage=" & TotalPages &"'>Last</a>"
	  end if
	  
	  %>
	  </b></font></td>
</tr>
</table>
</div>
</div>
<%
if TotalPages > 1 then
Response.Write "<center>"
Response.Write "<font face='arial' size='1' color=white>"
Response.Write "<b>Page:</b> "
Response.Write "<select name='whatever' onChange='window.location=this.options[this.selectedIndex].value'>"
Response.Write "<option value='view.asp?msgID="&id&"&seepage="&seepage&" SELECTED'> - "&seepage&" - </option>"
for i = 1 to TotalPages
if cStr(seepage) <> cStr(i) then
Response.Write "<option value='view.asp?msgID="&id&"&seepage="&i&"'> - "&i&" - </option>"
end if
next
Response.Write "</select>"
Response.Write "</font>"
Response.Write "</center>"

  Response.Write "<font face=arial size=1 color=white>"
  Response.Write "<p align=center>"
  Response.Write("Page " & seepage & " of " & TotalPages & "</p>")
  Response.Write "</font>"
end if
  %>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
