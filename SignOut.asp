<% 
    Dim LoggedInDate
    Dim LoggedOutDate

    LoggedInDate = Session("SLoggedDT")
    LoggedOutDate = Now()

    'response.write("User was Logged On: " & FormatDateTime(LoggedInDate,1) & " - " & FormatDateTime(LoggedInDate,3))
    'response.write("User was Logged Off: " & FormatDateTime(LoggedOutDate,1) & " - " & FormatDateTime(LoggedOutDate,3))

    Dim SDuration

    SDuration = DateDiff("s",LoggedInDate,LoggedOutDate)

    'response.write(SDuration)

    Function Duration(SecDuration)
        Dim TotalDuration
        TotalDuration = DateAdd("s", SecDuration, #00:00:00#)

        if SecDuration < 3600 then
        'response.write("<br> Total Duration : " & "00:" & Minute(TotalDuration) & ":" & Second(TotalDuration) )
        Duration = "00:" & Minute(TotalDuration) & ":" & Second(TotalDuration)
        else
        'response.write("<br> Total Duration : " &  Hour(TotalDuration) & ":" & Minute(TotalDuration) & ":" & Second(TotalDuration) )
        Duration = Hour(TotalDuration) & ":" & Minute(TotalDuration) & ":" & Second(TotalDuration)
        end if
    end Function

    Session.abandon
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <link rel="stylesheet" href="CSS/StyleHeader.css">
    <title>Home Expense Tracker - Sign Out</title>
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
            height: 40px;
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
    <div class="container-fluid">
        <br>
        <br>

        <div class="row">
            <div class="col">
                <img src="Images/Banner.png" alt="Banner" style="width: 100%; height: 350px;">
            </div>
        </div>

        <br>

        <div class="row">
            <div class="col-md-12">
                <!-------------------------Navigation Bar----------------------->
                <nav class="NavBar">
                    <ul>
                        <li>
                            <a href="Login.asp">Login Again</a>
                        </li>
                    </ul>
                </nav>
            </div>
        </div>

        <br>

        <div class="row text-center bg-success">
            <div class="col">
                <h3><% response.write(Session("StxtUserName"))%> You Have been Logged Out Successfully</h3>
            </div>
        </div>

        <br><br>

        <div class="row text-center">
            <div class="col-4">
                <h5>Login</h5>
            </div>
            <div class="col-4">
                <h5>Logout</h5>
            </div>
            <div class="col-4">
                <h5>Duration</h5>
            </div>
        </div>
        <div class="row text-center" style="background-color: lightgrey;">
            <div class="col-4">
                <h5><% response.write(FormatDateTime(LoggedInDate,1) & " - " & FormatDateTime(LoggedInDate,3)) %></h5>
            </div>
            <div class="col-4">
                <h5><% response.write(FormatDateTime(LoggedOutDate,1) & " - " & FormatDateTime(LoggedOutDate,3)) %></h5>
            </div>
            <div class="col-4">
                <h5><% response.write(Duration(SDuration)) %></h5>
            </div>
        </div>

        <br><br>
        </div>

        <footer style="background-color: rgb(54, 54, 54); color: whitesmoke;">
            <div class="container-fluid">
                <div class="row justify-content-center" style="height: 80px;">
                    <div class="col-3" style="padding: 20px;">
                        <h5>About Us</h5>
                        <span>This website is created by Nabeel Sohail.</span>
                    </div>
    
                    <div class="vl"></div>
    
                    <div class="col-5 text-center" style="padding: 35px;">
                        <ul class="links">
                            <li>
                                <a href="Login.asp">Login Again</a>
                            </li>
                        </ul>
                    </div>
    
                    <div class="vl"></div>
    
                    <div class="col-3 text-center" style="padding: 20px;">
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
                    <div class="col text-center" style="font-size: small; padding: 10px;">
                        Copyright &copy; 2020 - <% response.write(Year(Date()))%>, My Home Expense Tracker. All Rights
                        Reserved
                    </div>
                </div>
            </div>
        </footer>
    </div>
</body>

</html>