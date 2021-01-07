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
    Dim RSTransaction
    'Dim CS
    Dim RSCount

    Set Conn = Server.CreateObject("ADODB.Connection")
    Set RSTransaction = Server.CreateObject("ADODB.RecordSet")
    Set RSCount = Server.CreateObject("ADODB.RecordSet")

    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS

    Dim TableName
    Dim RecPerPage
    Dim RecNumber
    Dim PageNumber
    Dim SkipRec
    Dim LastPage

    

    If Request.QueryString("QsPageNumber")="" then
        PageNumber = 1
        SkipRec=0
    else
        PageNumber = Cint(request.QueryString("QsPageNumber"))
        SkipRec = (PageNumber*RecPerPage)-RecPerPage
    End if
    
    
    
    'If request.Form("AccTableName")="" then
    '    TableName = "HomeAccount"
    'else
    '    TableName = request.Form("AccTableName")
    'end if



    'Dim QryStr
    'QryStr = "SELECT dbo.HomeAccount.ID, dbo.HomeAccount.TransactionDate, dbo.HomeAccount.Description, dbo.HomeAccount.Credit, dbo.HomeAccount.Debit, dbo.HomeAccount.Balance, dbo.HomeAccount.CategoryID, " & _
     '         "dbo.Categories.Category, dbo.HomeAccount.PersonID, dbo.Persons.PersonName " & _ 
      '        "FROM dbo.Categories INNER JOIN " & _
       '       "dbo.HomeAccount ON dbo.Categories.CategoryID = dbo.HomeAccount.CategoryID INNER JOIN " & _
        '      "dbo.Persons ON dbo.HomeAccount.PersonID = dbo.Persons.PersonID ORDER BY HomeAccount.ID desc"  
    

    'Note: Request.Form("ObjectName") reads data after form submission/posted
    'Note: Request.querystring("ObjectName") Reads data from URL Query String Variable
    'Note: Request("ObjectName") reads data from either Form Object or Query String Object where Found

    
    'Using Form Object Variable
    'If request.Form("AccTableName")="" then
    '    QryStr = "SELECT * FROM HomeAccount WHERE (ID < 0)"
    'else
    '    TableName = request.Form("AccTableName")

'        QryStr = "SELECT " & TableName & ".ID, " & TableName & ".TransactionDate, " & TableName & ".Description, " & TableName & ".Credit, " & TableName & ".Debit, " & TableName & ".Balance, " & _
'                  TableName & ".CategoryID, dbo.Categories.Category, " & TableName & ".PersonID, dbo.Persons.PersonName " & _ 
'                  "FROM dbo.Categories INNER JOIN " & _
'                  TableName & " ON dbo.Categories.CategoryID = " & TableName & ".CategoryID INNER JOIN " & _
'                  "dbo.Persons ON " & TableName & ".PersonID = dbo.Persons.PersonID ORDER BY " & TableName & ".ID desc"         
'    End if      
    
    'Using Request("ObjectName")
    If request("AccTableName")="" then
        QryStr = "SELECT * FROM HomeAccount WHERE (ID < 0)"
    else
        TableName = request("AccTableName")

        QryStr = "SELECT " & TableName & ".ID, " & TableName & ".TransactionDate, " & TableName & ".Description, " & TableName & ".Credit, " & TableName & ".Debit, " & TableName & ".Balance, " & _
                  TableName & ".CategoryID, dbo.Categories.Category, " & TableName & ".PersonID, dbo.Persons.PersonName " & _ 
                  "FROM dbo.Categories INNER JOIN " & _
                  TableName & " ON dbo.Categories.CategoryID = " & TableName & ".CategoryID INNER JOIN " & _
                  "dbo.Persons ON " & TableName & ".PersonID = dbo.Persons.PersonID ORDER BY " & TableName & ".ID desc"         
    End if 

    RSTransaction.Open QRYSTR,Conn

    If request("AccTableName")<>"" then
        RSCount.Open "SELECT COUNT(ID) AS TotalRecords FROM " & TableName,conn

        TotalRec = RSCount("TotalRecords")
        RecPerPage=15

        If RSCount.EOF  or RSCount("TotalRecords")=1 then
            LastPage = 0
        else
            LastPage = Cstr((RSCount("TotalRecords")/RecPerPage))

            If InStr(LastPage,".") > 1 then
                LastPage = cint(LEFT(LastPage,InStr(LastPage,".")-1)) + 1
            end if
        End If
    else
        LastPage=0
    end if
    
    
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <link rel="stylesheet" href="CSS/StyleViewTrans.css">
    <title>Home Expense Tracker - Transactions</title>
    <style>
        /* Select Table Form */

        .wrapper {
            border: 2px solid black;
            background-color: lightgrey;
            width: 20%;
            margin: 0px auto;
            text-align: center;
            height: max-content;
            padding: 10px;
            border-radius: 4px;
        }

        .wrapper label {
            font-size: 16px;
        }

        .combo-account {
            align-content: center;
            font-weight: 700;
        }

        .select-account {
            border: 1px solid black;
            border-radius: 10px;
            padding: 5px;
            font-weight: 700;
            width: 300px;
        }

        .btn-select {
            background-color: rgb(165, 178, 192);
            border: 1px solid black;
            border-radius: 10px;
            text-align: center;
            padding: 5px;
            font-weight: 700;
            width: 15%;
            cursor: pointer;
            width: 50%;
        }

        .btn-select:hover {
            color: white;
            border: 1px solid whitesmoke;
        }

        .form-control {
            width: 80%;
            margin: auto;
        }

        .add-new {
            margin: 0px auto;
            text-align: center;
            padding: 10px;
            width: 15%;
        }

        .btn-addnew {
            background-color: rgb(165, 178, 192);
            border: 1px solid black;
            border-radius: 10px;
            padding: 8px;
        }

        .table-wrapper {
            background-color: white;
            margin: 10px;
            margin-top: 30px;
        }

        .table-wrapper th {
            text-align: center;
        }

        .icon {
            width: 25px;
            height: 25px;
        }

        .page-nav {
            width: 20%;
            margin: 0px auto;
            text-align: center;
            height: max-content;
            padding: 10px;
        }

        .btn-primary {
            padding: 8px 20px;
        }
    </style>
</head>

<body>
    <!--- #include file="Header.asp" -->

    <div class="container-fluid">

        <div class="row">
            <div class="col">
                <div class="header">

                    <h1>View Transactions</h1>
                    <%
                            If request("AccTableName")<>"" then
                                
                                Dim RSAccounts
                                Set RSAccounts = Server.CreateObject("ADODB.RecordSet")
            
                                RSAccounts.Open "SELECT TableDescription FROM AccountsTable WHERE (TableName = '" & TableName & "')", Conn
            
                                Response.Write("<h2>" & RSAccounts("TableDescription") & "</h2>")
            
                                RSAccounts.close
                                Set RSAccounts = Nothing
                            end if
                        %>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="bg-warning text-center">
                    <h3>
                        <% response.Write(Session("ErrorQsID")) %>
                        <% response.Write(Session("ErrorAction")) %>
                        <% response.Write(Session("ErrorTableName")) %>
                        <% response.Write(Session("ErrorFound")) %>
                        <% response.Write(Session("DateError")) %>
                        <% response.Write(Session("CreditError")) %>
                        <% response.Write(Session("DebitError")) %>
                    </h3>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="wrapper">
                    <form action="ViewTransaction.asp" method="POST" class="combo-account">
                        <div class="form-group">
                            <label for="AccTableName">Select Account</label>
                            <select name="AccTableName" class="select-account form-control">
                                <%
                        Set RSAccounts = Server.CreateObject("ADODB.RecordSet")
                    
                        RSAccounts.Open "SELECT TableName, TableDescription FROM AccountsTable",Conn
            
                        do while Not RSAccounts.EOF
                      
                        if RSAccounts("TableName") = TableName then  
                    %>
                                <option value="<% response.write(RSAccounts("TableName")) %>" selected>
                                    <% response.write(RSAccounts("TableDescription")) %></option>
                                <%
                        else %>
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
                            <br>
                            <input type="submit" value="Select Account" class="btn-select">
                        </div>
                    </form>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="add-new">
                    <div class="btn-addnew">
                        <a href="NewTransaction.asp?AccTableName=<% response.write(TableName) %>" class="">
                            <img src="images/add.png" alt="Add New" class="icon">
                            <span style="color: black;">Add New Transaction</span>
                        </a>
                    </div>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="table-wrapper">
                    <table class="table table-bordered table-hover">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 10%;">Transaction Id</th>
                                <th style="width: 10%;">Transaction Date</th>
                                <th style="width: 10%;">Category</th>
                                <th style="width: 10%;">Person</th>
                                <th style="width: 30%;">Description</th>
                                <th style="width: 10%;">Credit</th>
                                <th style="width: 10%;">Debit</th>
                                <th style="width: 10%;">Balance</th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                                Dim SkipCounter
                                SkipCounter=1
                                RecNumber = 0
            
                                'response.Write("Skip Rec =" & SkipRec)
                                'response.Write("<br>Page Number =" & PageNumber)
            
                                do while RSTransaction.eof=false
                                
                                if SkipCounter > SkipRec then    
                            %>
                            <tr>
                                <td><% response.Write(RSTransaction("Id")) %></td>
                                <td><% response.Write(RSTransaction("TransactionDate")) %></td>
                                <td><% response.Write(RSTransaction("Category")) %></td>
                                <td><% response.Write(RSTransaction("PersonName")) %></td>
                                <td><% response.Write(RSTransaction("Description")) %></td>
                                <td><% response.Write("PKR " & FormatNumber(RSTransaction("Credit"))) %></td>
                                <td><% response.Write("PKR " & FormatNumber(RSTransaction("Debit"))) %></td>
                                <td><% response.Write("PKR " & FormatNumber(RSTransaction("Balance"))) %></td>
                                <td><a
                                        href="EditTransaction.asp?QsID=<% response.Write(RSTransaction("Id")) %>&QSTableName=<% Response.write(TableName) %>&Action=1"><img
                                            src="images/edit.png" alt="Edit" class="icon"></a></td>
                                <td><a
                                        href="EditTransaction.asp?QsID=<% response.Write(RSTransaction("Id")) %>&QSTableName=<% Response.write(TableName) %>&Action=2"><img
                                            src="images/delete.png" alt="Delete" class="icon"></a></td>
                            </tr>
                            <%
                                    'response.Write("Rec Num = " & RecNumber)
                                    RecNumber = RecNumber + 1
                                
                                End if
            
                                If RecPerPage = RecNumber then
                                    'PageNumber = PageNumber+1
                                    exit do
                                end if
                                
                                SkipCounter = SkipCounter+1
                               RSTransaction.MoveNext
                               loop 
            
                                RSTransaction.Close
                                Set RSTransaction = Nothing
            
                                Conn.Close
                                Set Conn = Nothing
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="page-nav">
                    <% if LastPage = 0 or PageNumber <=1 then %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=1"
                        class="btn btn-primary disabled">First</a>
                    <% else %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=1"
                        class="btn btn-primary">First</a>
                    <% End if %>

                    <% if pagenumber > 1 then %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=<% response.write(PageNumber-1) %>"
                        class="btn btn-primary">Previous</a>
                    <% else %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=<% response.write(PageNumber-1) %>"
                        class="btn btn-primary disabled">Previous</a>
                    <% End if %>

                    <% if LastPage > 1 then %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=<% response.write(PageNumber+1) %>"
                        class="btn btn-primary">Next</a>
                    <% else %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=<% response.write(PageNumber+1) %>"
                        class="btn btn-primary disabled">Next</a>
                    <% end if %>

                    <% if LastPage >1 then %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=<% response.write(LastPage) %>"
                        class="btn btn-primary">Last</a>
                    <% else %>
                    <a href="ViewTransaction.asp?AccTableName=<% response.write(TableName) %>&QsPageNumber=<% response.write(LastPage) %>"
                        class="btn btn-primary disabled">Last</a>
                    <% End if %>
                </div>
            </div>
        </div>

    </div>

    <!--- #include file="Footer.asp" -->
</body>

</html>