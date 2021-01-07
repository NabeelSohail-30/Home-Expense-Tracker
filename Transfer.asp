<%
Session("STimeoutError")=""

if Session("STxtUserEmail")="" then
    Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
    response.Redirect("Login.asp")
end if
%>
<!--#include file="ValidateLogin.asp"-->
<%

%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Home Expense Tracker - Transfer Detail</title>
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <link rel="stylesheet" href="CSS/StyleNewTrans.css">
</head>

<body>
    <!--- #include file="Header.asp" -->
    <div class="container-fluid">
        <div class="row">
            <div class="col">
                <div class="wrapper">
                    <header>
                        <h1>Transfer Detail</h1>
                    </header>

                    <hr>

                    <form action="NewTransfer.asp" method="POST" class="combo-account">
                        <div class="form-group">
                            <label for="FormAccountFrom">From</label>
                            <select name="FormAccountFrom" class="select-account form-control">
                                <%
                                Dim RSAccounts
                                Set RSAccounts = Server.CreateObject("ADODB.RecordSet")
                    
                                RSAccounts.Open "SELECT TableName, TableDescription FROM AccountsTable",Conn
            
                                do while Not RSAccounts.EOF
                      
                            %>
                                <option value="<% response.write(RSAccounts("TableName")) %>"><% response.write(RSAccounts("TableDescription")) %></option>
                            <%
            
                                RSAccounts.MoveNext
                                Loop
            
                                RSAccounts.Close
                                Set RSAccounts = Nothing
                            %>
                            </select>
                        </div>

                        <div class="form-group">
                            <label for="FormAccountTo">To</label>
                            <select name="FormAccountTo" class="select-account form-control">
                                <%
                                Set RSAccounts = Server.CreateObject("ADODB.RecordSet")
                    
                                RSAccounts.Open "SELECT TableName, TableDescription FROM AccountsTable",Conn
            
                                do while Not RSAccounts.EOF
                      
                            %>
                                <option value="<% response.write(RSAccounts("TableName")) %>"><% response.write(RSAccounts("TableDescription")) %></option>
                            <%
            
                                RSAccounts.MoveNext
                                Loop
            
                                RSAccounts.Close
                                Set RSAccounts = Nothing
                            %>
                            </select>
                        </div>

                        <hr>

                        <div class="form-group">
                            <label for="">Transfer Date</label>
                            <input type="date" class="form-control" name="TrDate">
                            <span
                                style="color: red; font-size:medium;"><% response.Write(Session("DateError")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">Transfer Description</label>
                            <textarea cols="30" rows="3" class="form-control" name="TrDesc"></textarea>
                        </div>

                        <div class="form-group">
                            <label for="">Transfer Amount</label>
                            <input type="text" class="form-control" name="CreditAmount">
                            <span
                                style="color: red; font-size:medium;"><% response.Write(Session("CreditError")) %></span>
                        </div>

                        <div>
                            <input type="submit" value="Save Record" class="btn-select">
                        </div>
                        <div>
                            <span
                                style="color: red; font-size:medium;"><% response.Write(Session("ErrorFound")) %></span>
                        </div>
                    </form>

                </div>
            </div>
        </div>
    </div>
    <!--- #include file="Footer.asp" -->
</body>

</html>