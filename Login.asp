<%
    Dim CKUser
    CKUser=request.cookies("CookieUserName")
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Login Page</title>
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <link rel="stylesheet" href="CSS/StyleLogin.css">
</head>

<body>
    <div class="wrapper">
        <div class="header">
            <div class="top">
                <div class="logo">
                    <img src="Images/LoginAvatar.png" alt="Login">
                </div>
                <form action="ValidateLogin.asp" method="post">
                    <span style="color: red; font-size:medium;"><% response.Write(Session("SErrorEmail")) %></span>
                    <div class="input_field">
                        <% if CKUser <> "" then %>
                        <input type="email" placeholder="User Email" class="input" name="TxtUserEmail" value="<% response.write(CKUser) %>">
                        <% else %>
                        <input type="email" placeholder="User Email" class="input" name="TxtUserEmail" value="<% response.write(Session("STxtUserEmail")) %>">
                        <% end if %>
                    </div>
                    
                    <span style="color: red; font-size:medium;"><% response.Write(Session("SErrorPass")) %></span>    
                    <div class="input_field">
                        <input type="password" placeholder="Password" class="input" name="TxtPassword">
                    </div>

                    <span style="color: red; font-size:medium;"><% response.Write(Session("SErrorInvalid")) %></span>
                    <span style="color: red; font-size:medium;"><% response.Write(Session("STimeoutError")) %></span>
                    <div>
                        <label for="RememberMe">Remember my User Name</label>
                        <input type="checkbox" name="RememberMe" value="1" checked>
                    </div>
                    <div class="btn">
                        <input type="submit" value="Log In" class="login_btn">
                    </div>
                </form>
            </div>
        </div>
    </div>
</body>

</html>
