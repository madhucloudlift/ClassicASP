<html>
<head>
<title> [ AspBB ] - RePost </title>
</head>

<body bgcolor="#006699">
<%
id = Request.QueryString("msgID")
set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 

Set rsget = Server.CreateObject("ADODB.Recordset")
rsget.Open "SELECT * FROM board WHERE ThreadID="&id&" AND ThreadP=0", conn
error = request.querystring("error")
select case cstr(error)
case "1"
Response.Write "<p align=center><font face=arial color=silver size=2><b>You have not fill out Topic, Name or Body.  Server cannot process unless these field(s) are fill.</b></font>"
case "2"
Response.Write "<p align=center><font face=arial color=silver size=2><b>Sorry, Scripting is not allow on this MessageBoard.  Please fix your Message.</b></font>"
case "3"
Response.Write "<p align=center><font face=arial color=silver size=2><b>Sorry, you have enter an invalid e-mail address.  Please fix your e-mail address.</b></font>"
case "4"
Response.Write "<p align=center><font face=arial color=silver size=2><b>Sorry, you have enter an invalid URL. Please fix your URL.</b></font>"
case Else
Response.Write ""
end select
%>
<br>
<div align="center">
<table border="0" width="50%">
<tr>
<td color="blue">
<form method="post" action="resubmit.asp">
<font face="arial" color="white" size="2"><b>Topic: </b></font><font face="arial" color="white" size="2"><b>Re: <%=rsget("topic")%></b></font><br>
<input type="hidden" name="topic" value="Re: <%=rsget("topic")%>">
<input type="hidden" name="msgID" value="<%=id%>">
<font face="arial" color="white" size="2"><b>Name: </b></font><input type="text" name="name" maxlength="15"><br>
<font face="arial" color="white" size="2"><b>E-mail: </b></font><input type="text" name="email" size=25><br>
<font face="arial" color="white" size="2"><b>Your Website: </b></font><input type="text" name="http" size=30><br><br>
<font face="arial" color="white" size="2"><b>Message: </b></font><br>
<textarea name="body" rows="9" cols="55"></textarea><br>
<p align=center><input type="submit" value="Submit">
<input type="reset" value="Reset"></p>
</form>
</td>
</tr>
</table>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
