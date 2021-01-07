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
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <title>Home Expense Tracker - New Account</title>
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
                        <h1>New Account</h1>
                    </header>

                    <hr>

                    <form action="AddAccount.asp" method="POST" class="combo-account">
                        <div class="form-group">
                            <label for="">Account Name</label>
                            <input type="text" name="AccName" class="form-control">
                        </div>

                        <div class="form-group">
                            <label for="">Account Description</label>
                            <input type="text" name="AccDescription" class="form-control">
                        </div>

                        <div class="form-group">
                            <label for="">Opening Balance Date</label>
                            <input type="date" name="OpnBalDate" id="" class="form-control">
                        </div>

                        <div class="form-group">
                            <label for="">Opening Balance</label>
                            <input type="text" name="OpnBalance" id="" class="form-control">
                        </div>

                        <div class="form-group">
                            <input type="submit" value="Add New Account" class="btn-add">
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
    <!--- #include file="footer.asp" -->
</body>

</html>