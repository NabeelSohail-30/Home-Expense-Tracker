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
    <title>Home Expense Tracker - Sign Out</title>
</head>

<body>
    <!--- #include file="Header.asp" -->
    <div class="container-fluid">
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
    </div>
    <!--- #include file="footer.asp" -->
</body>

</html>