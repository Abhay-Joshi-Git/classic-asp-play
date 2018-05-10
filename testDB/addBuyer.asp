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


Status = -1
submitted = False

' Note - we need to include adovbs.inc so that we can access constants like adCmdStoredProc, adParamReturnValue
%>
<!-- #INCLUDE VIRTUAL = "Header.asp" -->
<!-- #INCLUDE VIRTUAL = "adovbs.inc" -->
<link type="text/css" rel="stylesheet" href="bootstrap.min.css" />
<link type="text/css" rel="stylesheet" href="main.css" />

<%

 on error resume next
Set cnnTest = Server.CreateObject("ADODB.Connection")
cnnTest.ConnectionString = Application("connectionString")
cnnTest.Open
    
Randomize
id = CStr(Year(Now()) + Month(Now()) + Day(Now())  + Hour(Now()) + Minute(Now()) + Second(Now())) & CStr(CInt(Rnd * 10))
id = CInt(id)


If Request.Form("name") <> "" Then
    submitted = true
    Set addUserCmd = Server.CreateObject("ADODB.Command")
    With addUserCmd
	    .ActiveConnection = cnnTest
	    .CommandType = adCmdStoredProc
	    .Prepared = True
	    .CommandText = "add_buyer"

	    .Parameters.Append .CreateParameter ("RC", adInteger, adParamReturnValue, 4) 
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput, 4, id)
		.Parameters.Append .CreateParameter("name", adVarchar, adParamInput, 50, Request.Form("name"))
        .Parameters.Append .CreateParameter("address", adVarchar, adParamInput, 50, Request.Form("address"))

        .Execute

        If err.number <> 0 Then
            response.Write("Error - " & err.Description)
        End If

         Status = .Parameters("RC")
    End With

    response.Write("<br />")

    Set addUserCmd = Nothing
End If
    
Set cnnTest = Nothing

%>

<div class="container-half container">
    <% If submitted = True And Status  = 0 Then%>
    <div class="card bg-success text-white" style="box-shadow: 0 0 10px -5px lightgrey; padding: 5px;">
        <h2>Buyer <% request.Form("Name") %> has been added sucessfully.</h2>
    </div>
    <br />
    <h3>Add Another Buyer :</h3>
    <% Else %>
    <h3>Add Buyer :</h3>
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