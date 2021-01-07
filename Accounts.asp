<%
    'Start - Validate logged user Session Timeout. If Session is timeout then it will redirect to te login page.
        Session("STimeoutError")=""

        if Session("STxtUserEmail")="" then
            Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
            response.Redirect("Login.asp")
        end if
    'end
%>
    <!--Start - Validate the Logged User Access Everytime on any page access.-->
    <!--Use of Include File method will save time to rewrite the code on every page.-->
        <!--#include file="ValidateLogin.asp"-->
    <!--End-->
<%
    'Database and RS Start 
        'Dim Conn 
        Dim RSAccount
        'Dim CS

        'Set Conn = Server.CreateObject("ADODB.Connection")
        Set RSAccount = Server.CreateObject("ADODB.RecordSet")
        
        'CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        'Conn.Open CS

        RSAccount.Open "SELECT * FROM AccountsTable",Conn
    'Database and RS end
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <title>Home Expense Tracker - Accounts</title>
    <style>
        .table-wrapper {
            background-color: white;
            margin: 10px auto;
            margin-top: 30px;
            width: 40%;

        }

        .table-wrapper th {
            text-align: center;
        }

        .icon {
            width: 25px;
            height: 25px;
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
    </style>
</head>

<body>
    <!--- #include file="Header.asp" -->
    <div class="container-fluid">
        <div class="row">
            <div class="col">
                <div class="add-new">
                    <div class="btn-addnew">
                        <a href="NewAccount.asp" class="">
                            <img src="images/add.png" alt="Add New" class="icon">
                            <span style="color: black;">Add New Account</span>
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
                                <th style="width: 10%;">Table Name</th>
                                <th style="width: 12%;">Table Description</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                                do while not RSAccount.EOF
                            %>
                            <tr>
                                <td><% response.write(RSAccount("TableName")) %></td>
                                <td><% response.write(RSAccount("TableDescription")) %></td>
                            </tr>
                            <%
                                RSAccount.MoveNext
                                loop 
            
                                RSAccount.Close
                                Set RSAccount = Nothing
            
                                Conn.Close
                                Set Conn = Nothing
                            %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <!--- #include file="footer.asp" -->
</body>

</html>