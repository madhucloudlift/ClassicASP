<%
response.buffer = true

topic = Replace(Server.HTMLEncode(Request.Form("topic")),"'","''")
name = Replace(Server.HTMLEncode(Request.Form("name")),"'","''")
email = Replace(Server.HTMLEncode(Request.Form("email")),"'","''")
http = Replace(Server.HTMLEncode(Request.Form("http")),"'", "''")
body =  Request.Form("body")
body = Replace(body , "'", "''")
body = Replace(body, vbcrlf, "<br>")

if topic = "" OR name = "" OR body = "" then
Response.Redirect("post.asp?error=1")
end if
if instr(LCase(body), "<script") <> 0 then 
Response.Redirect("post.asp?error=2")
end if
if email <> "" then 
  if instr(email, "@") = 0 OR instr(email, ".") = 0 then
  Response.Redirect("post.asp?error=3")
  end if
end if
if http <> "" then
  if instr(http, "http://") = 0 OR instr(http, ".") = 0 then
  Response.Redirect("post.asp?error=4")
  end if
end if

set conn = Server.CreateObject("ADODB.connection")
sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
"Data Source=" & Server.MapPath("\dangerduo\db\betaboard.mdb") & ";" & _ 
"Persist Security Info=False" 
conn.Open(sConnection) 

sqlstring = "INSERT INTO board ( [date], topic, name, email, http, body )  " &_ 
 "values ( # " &now()& " #, ' " &topic& " ' ,  ' " &name& " ' ,  ' " &email& " ', ' " &http& " ' , ' " &body& " ')" 
conn.Execute(sqlstring)
conn.Close 
set conn = Nothing
Response.redirect "default.asp"
%>