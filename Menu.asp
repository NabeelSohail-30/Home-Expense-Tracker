<%
    Session("STimeoutError")=""

    if Session("STxtUserEmail")="" then
        Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
        response.Redirect("Login.asp")
    end if
%>
    <!--#include file="ValidateLogin.asp"-->
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <link rel="stylesheet" href="CSS/Stylemenu.css">
    <title>Home Expense Tracker - Main Menu</title>
</head>

<body>
    <!--- #include file="Header.asp" -->
    <div class="container-fluid">

        <div class="row">
            <div class="col bg-success text-center">
                <% response.write(request.querystring("msg")) %>
            </div>
        </div>

        <br>

        <div class="container">
            <div class="row justify-content-center">
                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/Cash-512.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Transactions</h4>
                            <div class="card-text">Add, View, Edit, Delete Transactions </div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>

                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/accounting.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Accounts</h4>
                            <div class="card-text">Add, View, Edit, Delete Accounts </div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>

                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/list.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Categories</h4>
                            <div class="card-text">Add, View, Edit, Delete Categories</div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>

                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/audience.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Persons</h4>
                            <div class="card-text">Add, View, Edit, Delete Persons</div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>

                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/profile.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Manage Users</h4>
                            <div class="card-text">Add, View, Edit, Delete Users </div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>

                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/password.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Change Password</h4>
                            <div class="card-text">Change Password</div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>

                <div class="col-4 mt-3">
                    <div class="card text-center" style="border: 1px solid black;">
                        <img class="card-img-top" src="Images/logout.png" alt="Card image"
                            style="width: 250px; height: 250px; margin: auto;">

                        <hr>

                        <div class="card-body">
                            <h4 class="card-title">Sign Out</h4>
                            <div class="card-text">Click Here to Sign Out</div>
                            <br>
                            <a href="#" class="btn btn-primary stretched-link">Click Here</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!--- #include file="footer.asp" -->

</body>

</html>