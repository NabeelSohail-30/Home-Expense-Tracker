<!--#include file="MyFunc.asp"-->
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CSS/StyleHeader.css">
    <title>Header</title>
    <style>
        .btn-signout{
            display: inline-block;
            background-color: lightgray;
            color: black;
            padding: 5px;
        }

        .btn-signout:hover{
            display: inline-block;
            background-color: rgb(92, 92, 92);
            color: whitesmoke;
            padding: 5px;
        }
        body{
            margin-top: 5px;
        }
    </style>
</head>

<body>
    <div class="container-fluid">
        <div class="row align-content-center" style="height:40px;">
            <div class="col-7 align-self-center text-left">
                <h5>
                    Welcome <% response.write(Session("StxtUserName") & " (" & Session("STxtUserEmail") & ") - Active Users " & Application("visitors") ) %>
                </h5>
            </div>
            <div class="col-5 align-self-center text-right">
                <h5>
                    Logged On:
                    <% response.write(FormatDateTime(Session("SLoggedDT"),1) & " - " & FormatDateTime(Session("SLoggedDT"),3)) %>
                    <a href="SignOut.asp" class="btn-signout">Sign Out</a>
                </h5>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <img src="Images/Banner.png" alt="Banner" style="width: 100%; height: 350px;">
            </div>
        </div>

        <div class="row">
            <div class="col-md-12">
                <!-------------------------Navigation Bar----------------------->
                <nav class="NavBar">
                    <ul>
                        <li>
                            <a href="Menu.asp">Main Menu</a>
                        </li>

                        <li>
                            <a href="ViewTransaction.asp">Transactions</a>
                        </li>

                        <li>
                            <a href="Accounts.asp">Accounts</a>
                        </li>

                        <li>
                            <a href="Transfer.asp">Transfer</a>
                        </li>

                        <li>
                            <a href="Categories.asp">Categories</a>
                        </li>

                        <li>
                            <a href="Persons.asp">Persons</a>
                        </li>

                        <li>

                            <% if Session("SIsAdmin") <> "False" then %>
                            <a href="Users.asp">Manage Users</a>
                            <% else %>
                            <a href="#">Manage Users</a>
                            <% end if %>
                        </li>

                        <li>
                            <a href="#">Change Password</a>
                        </li>

                        <li>
                            <a href="SignOut.asp">Sign Out</a>
                        </li>
                    </ul>
                </nav>
            </div>
        </div>
    </div>
</body>

</html>