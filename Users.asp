<%
    if Session("SIsAdmin") = "False" then
        response.Redirect("Menu.asp")
    end if

    Session("STimeoutError")=""

    if Session("STxtUserEmail")="" then
        Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
        response.Redirect("Login.asp")
    end if

    Dim Conn 
    Dim RSLogin
    Dim CS

    Set Conn = Server.CreateObject("ADODB.Connection")
    Set RSLogin = Server.CreateObject("ADODB.RecordSet")
    
    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS

    RSLogin.Open "SELECT * FROM LoginDetails",Conn
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
            width: 50%;

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
                        <a href="NewUser.asp" class="">
                            <img src="images/add.png" alt="Add New" class="icon">
                            <span style="color: black;">Add New User</span>
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
                                <th style="width: 10%;">ID</th>
                                <th style="width: 15%;">User Full Name</th>
                                <th style="width: 15%;">User Email</th>
                                <th style="width: 5%;"></th>
                                <th style="width: 5%;"></th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                                do while not RSLogin.EOF
                            %>
                            <tr>
                                <td><% response.write(RSLogin("LoginId")) %></td>
                                <td><% response.write(RSLogin("UserFullName")) %></td>
                                <td><% response.write(RSLogin("UserEmail")) %></td>
                                <td><a href="NewUser.asp?QsID=<% response.Write(RSLogin("LoginId")) %>&Action=1"><img src="images/edit.png" alt="Delete" class="icon"></a></td>
                                <td><a href="DelUser.asp?QsID=<% response.Write(RSLogin("LoginId")) %>"><img src="images/delete.png" alt="Delete" class="icon"></a></td>
                            </tr>
                            <%
                                RSLogin.MoveNext
                                loop 
            
                                RSLogin.Close
                                Set RSLogin = Nothing
            
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