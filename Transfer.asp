<%
Session("STimeoutError")=""

if Session("STxtUserEmail")="" then
    Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
    response.Redirect("Login.asp")
end if
%>
<!--#include file="ValidateLogin.asp"-->
<%

    'Dim Conn 
    'Dim CS

    Set Conn = Server.CreateObject("ADODB.Connection")

    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS

    Dim TableName
    
    TableName = Request.QueryString("AccTableName")

%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Home Expense Tracker - New Transaction</title>
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
                        <h1>New Transaction</h1>
                    </header>

                    <hr>

                    <form action="SaveTransaction.asp" method="POST" class="combo-account">
                        <div>
                            <label for="AccTableName">Select Account</label>
                            <select name="AccTableName" class="select-account form-control">
                                <%
                                Dim RSAccounts
                                Set RSAccounts = Server.CreateObject("ADODB.RecordSet")
                    
                                RSAccounts.Open "SELECT TableName, TableDescription FROM AccountsTable",Conn
            
                                do while Not RSAccounts.EOF
                      
                                if RSAccounts("TableName") = TableName then  
                            %>
                                <option value="<% response.write(RSAccounts("TableName")) %>" selected>
                                    <% response.write(RSAccounts("TableDescription")) %></option>
                                <%
                            else 
                            %>
                                <option value="<% response.write(RSAccounts("TableName")) %>">
                                    <% response.write(RSAccounts("TableDescription")) %></option>
                                <%
                                end if
            
                                RSAccounts.MoveNext
                                Loop
            
                                RSAccounts.Close
                                Set RSAccounts = Nothing
                            %>
                            </select>
                        </div>

                        <hr>

                        <div class="form-group">
                            <label for="">Transaction Date</label>
                            <input type="date" class="form-control" name="TrDate">
                            <span
                                style="color: red; font-size:medium;"><% response.Write(Session("DateError")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">Category</label>
                            <select name="SelectCategory" class="select-account form-control">
                                <%
                                Dim RSCatg
                                Set RSCatg = Server.CreateObject("ADODB.RecordSet")
                    
                                RSCatg.Open "SELECT CategoryID, Category FROM Categories",Conn
            
                                do while Not RSCatg.EOF 
                            %>
                                <option value="<% response.write(RSCatg("CategoryID")) %>">
                                    <% response.write(RSCatg("Category")) %></option>
                                <%
                                RSCatg.MoveNext
                                Loop
            
                                RSCatg.Close
                                Set RSCatg = Nothing
                            %>
                            </select>
                        </div>

                        <div class="form-group">
                            <label for="">Person</label>
                            <select name="SelectPerson" class="select-account form-control">
                                <%
                                Dim RSPerson
                                Set RSPerson = Server.CreateObject("ADODB.RecordSet")
                    
                                RSPerson.Open "SELECT PersonID, PersonName FROM Persons",Conn
            
                                do while Not RSPerson.EOF 
                            %>
                                <option value="<% response.write(RSPerson("PersonID")) %>">
                                    <% response.write(RSPerson("PersonName")) %></option>
                                <%
                                RSPerson.MoveNext
                                Loop
            
                                RSPerson.Close
                                Set RSPerson = Nothing
                            %>
                            </select>
                        </div>

                        <div class="form-group">
                            <label for="">Transaction Description</label>
                            <textarea cols="30" rows="3" class="form-control" name="TrDesc"></textarea>
                        </div>

                        <div class="form-group">
                            <label for="">Credit Amount</label>
                            <input type="text" class="form-control" name="CreditAmount">
                            <span
                                style="color: red; font-size:medium;"><% response.Write(Session("CreditError")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">Debit Amount</label>
                            <input type="text" class="form-control" name="DebitAmount">
                            <span
                                style="color: red; font-size:medium;"><% response.Write(Session("DebitError")) %></span>
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