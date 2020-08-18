<html>
<head>
<title> [ AspBB ] - Post </title>
</head>
<body bgcolor="#006699">
<%
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
<form method="post" action="submit.asp">
<font face="arial" color="white" size="2"><b>Topic: </b></font><input type="text" name="topic"><br>
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
</body>
</html>
