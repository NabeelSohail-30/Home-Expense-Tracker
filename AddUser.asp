<%
'Session Timeout Start
Session("STimeoutError")=""

    if Session("STxtUserEmail")="" then
        Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
        response.Redirect("Login.asp")
    end if
%>
    <!--#include file="ValidateLogin.asp"-->
<%

'Variable Declaration Start
    Dim Name
    Dim Email
    Dim Password
    Dim ConfirmPassword

    'Dim ErrorFound

    'Dim Conn 
    'Dim CS
    'Variable Declaration End

    Session("ErrorName")=""
    Session("ErrorEmail")=""
    Session("ErrorPassword")=""

    'Initializing Variables Start
    Name = request.Form("UserName")
    Email = request.Form("UserEmail")
    password = request.Form("UserPass")
    ConfirmPassword = request.Form("ConfirmPass")
    ErrorFound=false
    'Initializing Variables End

    if Name = "" then
        Session("ErrorName") = "User Name cannot be NULL"
        ErrorFound=True
    end if

    if Email = "" then
        Session("ErrorEmail") = "User Email cannot be NULL"
        ErrorFound=True
    end if

    if Password = "" then
        Session("ErrorPassword") = "Password cannot be NULL"
        ErrorFound=True
    end if

    if ConfirmPassword = "" then
        Session("ErrorPassword") = "Confirm Password cannot be NULL"
        ErrorFound=True
    end if

    If Password <> ConfirmPassword then
        Session("ErrorPassword") = "Confirm Password not matched"
        ErrorFound=True
        response.Redirect("NewUser.asp")
    end if

    If ErrorFound=true then
        response.Redirect("NewUser.asp")
    end if

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'Adding New User Start
        'Dim QryStr

        QryStr = "INSERT INTO LoginDetails (UserFullName,UserEmail,Password) Values ('" & Name & "', '" & Email & "', '" & Password & "')"
    
        'response.Write(qrystr) 

        Conn.Execute QryStr
    'End

    Conn.Close
    set conn = nothing

    response.Redirect("Users.asp")
%>