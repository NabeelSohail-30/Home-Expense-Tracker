<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Footer</title>

    <style>
        * {
            margin: 0px;
            padding: 0px;
            box-sizing: border-box;
            outline: none;
            list-style: none;
            text-decoration: none;
        }

        .vl {
            border: 1px solid rgba(128, 128, 128, 0.411);
            height: 80px;
            margin: 20px 0;
        }

        .media-icons img {
            width: 25px;
            height: 25px;
        }

        .media-icons li {
            display: inline-block;
        }

        .media-icons a {
            display: block;
            padding: 0 6px;
        }

        .links li {
            display: inline-block;
        }

        .links a {
            display: block;
            padding: 0 6px;
            text-decoration: none;
            color: whitesmoke;
        }

        .links a:hover {
            display: block;
            padding: 0 6px;
            text-decoration: underline;
            background-color: whitesmoke;
            color: black;
        }
    </style>
</head>

<body>
    <footer style="background-color: rgb(54, 54, 54); color: whitesmoke;">
        <div class="container-fluid">
            <div class="row justify-content-center mt-3">
                <div class="col-3 pt-4">
                    <h5>About Us</h5>
                    <span>This website is created by Nabeel Sohail.</span>
                </div>

                <div class="vl"></div>

                <div class="col-5 text-center pt-4">
                    <h5>Quick Links</h5>
                    <ul class="links">
                        <li><a href="Menu.asp">Main Menu</a></li>
                        <li><a href="ViewTransaction.asp">Transactions</a></li>
                        <li><a href="AddAccount.asp">Accounts</a></li>
                        <li><a href="Transfer.asp">Transfer</a></li>
                        <li><a href="Categories.asp">Categories</a></li>
                        <li><a href="Persons.asp">Persons</a></li>
                        <li><a href="#">User Management</a></li>
                        <li><a href="#">Change Password</a></li>
                        <li><a href="s">Sign Out</a></li>
                    </ul>
                </div>

                <div class="vl"></div>

                <div class="col-3 text-center pt-4">
                    <h5>Social Media Links</h5>
                    <ul class="media-icons">
                        <li><a href="#"><img src="images/facebook.png" alt=""></a></li>
                        <li><a href="#"><img src="images/instagram.png" alt=""></a></li>
                        <li><a href="#"><img src="images/twitter.png" alt=""></a></li>
                        <li><a href="#"><img src="images/youtube.png" alt=""></a></li>
                        <li><a href="#"><img src="images/linkedin.png" alt=""></a></li>
                    </ul>
                </div>
            </div>

            <hr style="color:gray;background-color:gray">

            <div class="row" style="padding-bottom: 12px;">
                <div class="col text-center" style="font-size: small;">
                    Copyright &copy; 2020 - <% response.write(Year(Date()))%>, My Home Expense Tracker. All Rights
                    Reserved
                </div>
            </div>
        </div>
    </footer>
</body>

</html>