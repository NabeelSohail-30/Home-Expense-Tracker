<%
Session("STimeoutError")=""

if Session("STxtUserEmail")="" then
    Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
    response.Redirect("Login.asp")
end if
%>
<!--#include file="ValidateLogin.asp"-->
<%

    if request.QueryString("Action")="1" then
        'Dim Conn
        'Dim CS
        'Dim RSLogin

        'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
        'Opening Db End

        Set RSLogin = Server.CreateObject("ADODB.RecordSet")
        RSLogin.Open "SELECT * FROM LoginDetails WHERE(LoginId = " & request.QueryString("QsId") & ")",Conn
    end if
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <% If request.QueryString("Action")="1" then %>
    <title>Home Expense Tracker - Edit User</title>
    <% else %>
    <title>Home Expense Tracker - Add User</title>
    <% end if %>
    
    <style>
        * {
            margin: 0px;
            padding: 0px;
            box-sizing: border-box;
            outline: none;
            list-style: none;
            text-decoration: none;
        }

        body {
            background-color: lightslategray;
        }

        .wrapper {
            border: 2px solid black;
            background-color: lightgrey;
            width: 35%;
            margin: 50px auto;
            text-align: center;
            height: max-content;
            padding: 10px;
            border-radius: 4px;
        }

        .combo-account {
            align-content: center;
            font-weight: 700;
        }

        .btn-add {
            background-color: rgb(165, 178, 192);
            border: 1px solid black;
            border-radius: 10px;
            text-align: center;
            padding: 5px;
            font-weight: 700;
            width: 30%;
            cursor: pointer;
        }

        .btn-add:hover {
            color: white;
            border: 1px solid whitesmoke;
        }

        .form-control {
            width: 50%;
            margin: auto;
        }
    </style>
</head>

<body>
    <!--- #include file="Header.asp" -->
    <div class="container-fluid">
        <div class="row">
            <div class="col">
                <div class="wrapper">
                    <header>
                        <% If request.QueryString("Action")="1" then %>
                        <h1>Edit User</h1>
                        <% else %>
                        <h1>New User</h1>
                        <% end if %>
                    </header>

                    <hr>

                    <% If request.QueryString("Action")="1" then %>
                    <form action="EditUser.asp" method="POST" class="combo-account">
                        <input type="hidden" name="TxtId" value="<% response.write(RSLogin("LoginId")) %>"/>
                        <div class="form-group">
                            <label for="">User Full Name</label>
                            <input type="text" class="form-control" name="UserName" value="<% response.write(RSLogin("UserFullName")) %>">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorName")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">User Email</label>
                            <input type="email" class="form-control" name="UserEmail" value="<% response.write(RSLogin("UserEmail")) %>">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorEmail")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">Current Password</label>
                            <input type="text" class="form-control" name="UserPass" value="<% response.write(RSLogin("Password")) %>">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorPassword")) %></span>
                        </div>

                        <div>
                            <input type="submit" value="Edit User" class="btn-add">
                        </div>
                    </form>
                    <% else %>
                    <form action="AddUser.asp" method="POST" class="combo-account">
                        <div class="form-group">
                            <label for="">User Full Name</label>
                            <input type="text" class="form-control" name="UserName">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorName")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">User Email</label>
                            <input type="email" class="form-control" name="UserEmail">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorEmail")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">Enter Password</label>
                            <input type="password" class="form-control" name="UserPass">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorPassword")) %></span>
                        </div>

                        <div class="form-group">
                            <label for="">Confirm Password</label>
                            <input type="password" class="form-control" name="ConfirmPass">
                            <span style="color: red; font-size:medium;"><% response.write(Session("ErrorPassword")) %></span>
                        </div>

                        <div>
                            <input type="submit" value="Add User" class="btn-add">
                        </div>
                    </form>
                    <% end if %>
                        

                </div>
            </div>
        </div>
    </div>
    <!--- #include file="Footer.asp" -->
</body>

</html>