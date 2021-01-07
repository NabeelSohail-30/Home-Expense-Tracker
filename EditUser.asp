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
    Dim Id

    'Dim ErrorFound

    'Dim Conn 
    'Dim CS
    'Variable Declaration End

    'Initializing Variables Start
    Name = request.Form("UserName")
    Id = request.Form("TxtId")
    Email = request.Form("UserEmail")
    password = request.Form("UserPass")
    ErrorFound=false
    'Initializing Variables End

    'response.write (Id)
    'response.end

    Session("ErrorName")=""
    Session("ErrorEmail")=""
    Session("ErrorPassword")=""

    if Name = "" or isnull(name)=true then
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

    If ErrorFound=true then
        response.Redirect("NewUser.asp")
    end if

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'Edit User Start
        Set RSLogin = Server.CreateObject("ADODB.RecordSet")
        RSLogin.Open "SELECT * FROM LoginDetails WHERE(LoginId = " & id & ")",Conn

        QryStr = "UPDATE LoginDetails SET UserFullName = '" & Name & "',UserEmail = '" & Email & "',Password = '" & Password & "' WHERE (LoginId = " & id & ")"

        'response.write(qrystr)
        'response.end
        Conn.Execute QryStr
    'End

    Conn.Close
    set conn = nothing

    response.Redirect("Users.asp")
%>