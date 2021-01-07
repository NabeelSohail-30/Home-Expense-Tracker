<%
Session("STimeoutError")=""

if Session("STxtUserEmail")="" then
    Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
    response.Redirect("Login.asp")
end if
%>
<!--#include file="ValidateLogin.asp"-->
<%

    'Dim Conn 
    'Dim CS

    Set Conn = Server.CreateObject("ADODB.Connection")
    
    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS

    'Dim QryStr

    QryStr = "DELETE FROM LoginDetails WHERE(LoginId = " & cint(request.querystring("QsId")) & ")"

    'response.write QryStr

    Conn.execute QryStr
    
    Conn.close 
    Set Conn = Nothing

    response.redirect("Users.asp")

%>