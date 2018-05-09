<%@ Language=VBScript %>
<%
Option Explicit
Dim message
Dim cnnTest
Dim addUserCmd
Dim rs
Dim id
Dim submitted
Dim name
Dim address
Dim Status
Dim sql


    submitted = False
%>
<!-- #INCLUDE VIRTUAL = "Header.asp" -->
<link type="text/css" rel="stylesheet" href="bootstrap.min.css" />
<link type="text/css" rel="stylesheet" href="main.css" />

<%

on error resume next
Set cnnTest = Server.CreateObject("ADODB.Connection")
cnnTest.ConnectionString = Application("connectionString")
cnnTest.Open


response.Write("form count" & Request.Form.Count & "  " & Request.Form)
response.Write("<br />")
Randomize
id = (Month(Now()) & "" & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now()) &"" &  CInt(Rnd * 10)) + 2
response.Write(id)
response.Write(Request.Form("name"))

' replace this with stored proc
sql = "INSERT INTO dbo.Buyer VALUES ( " & id &", ' " & Request.Form("name") & "', ' " & Request.Form("address") & "')"

If Request.Form("name") Then
    response.Write("sql  - "  & sql)
    cnnTest.Execute(sql)
End If
    
Set cnnTest = Nothing

%>

<div class="container-half container">
    <% If submitted  Then%>
    <h3>Add Buyer :</h3>
    <% Else %>
    <h3>Add Another Buyer :</h3>
    <% End If %>
    <br />
    <form method="POST" action="addBuyer.asp">
        <div class="form-group">
            <label for="name">Name :</label>
            <input name="name" class="form-control"  id="name" type="text" placeholder="Name" required />
        </div>
        <div class="form-group">
            <label for="address">address :</label>
            <input name="address" class="form-control"  id="address" placeholder="Address"/>
        </div>
        <br />
        <input class="btn btn-primary" type="submit" value="Submit"/>
    </form> 
</div>