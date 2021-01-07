<%
    
    if Session("STxtUserEmail")="" and Session("STxtPassword")="" then
        dim ErrorFound
        ErrorFound=False

        Session("SErrorEmail")=""
        Session("SErrorPass")=""

        if request.form("TxtUserEmail")="" or isnull(request.form("TxtUserEmail"))=true then
            Session("SErrorEmail")="Please Enter Login Email"
            ErrorFound=true
        end if

        if request.form("TxtPassword")="" or isnull(request.form("TxtPassword"))=true then
            Session("SErrorPass")="Please Enter Login Password"
            Session("Stxtuseremail")=request.form("TxtUserEmail")
            ErrorFound=true
        end if

        if ErrorFound=true then
                RSLogin.Close 
                Set RSLogin = Nothing
                
                Conn.close
                Set Conn = Nothing
            response.Redirect("Login.asp")
        end if
    end if

    Dim Conn
    Dim RSLogin
    Dim CS
    
    Set Conn = Server.CreateObject("ADODB.Connection")
    Set RSLogin = Server.CreateObject("ADODB.RecordSet")
    
    'Method 1 Connect Using CS
    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS
    
    'Method 2 Connect using ODBC Name
    'CS = "DSN=MyHomeExpense"
    'conn.Open CS,"homeexpense","Nabeel30"

    Dim mTxtUserEmail
    Dim mTxtPassword
    
    if request.form("TxtUserEmail") <> "" and request.form("TxtPassword") <> "" then
        mTxtUserEmail = request.Form("TxtUserEmail")
        mTxtPassword = request.Form("TxtPassword")
    else
        mTxtUserEmail = Session("STxtUserEmail")
        mTxtPassword = Session("STxtPassword")
    end if

    'Session("STxtUserEmail")=""
    'Session("STxtUserName")=""
    'Session("STxtPassword")=""
    
    Session("SErrorInvalid")=""
    Session("STimeoutError")=""

    Dim QryStr

    QryStr = "SELECT * FROM LoginDetails WHERE (UserEmail = '" &  mTxtUserEmail & "') AND (Password = '" & mTxtPassword & "')"

    'response.Write(qrystr)
    'response.end

    rslogin.Open qrystr,conn

    if rslogin.EOF=true then
        Session("SErrorInvalid")=("Invalid UserName or Password")
        Session("Stxtuseremail")=mTxtUserEmail

        Session("STxtUserName")=""
        Session("STxtPassword")=""

        response.Redirect("Login.asp")
    else
        if RSLogin("Active")= 0 then
            Session("SErrorInvalid")="Your Account is not Active anymore. Please Contact System Administrator"
            Session("Stxtuseremail")=mTxtUserEmail
            Session("STxtUserName")=""
            Session("STxtPassword")=""

            response.Redirect("Login.asp")
        else
        
            if request.form("TxtUserEmail") <> "" and request.form("TxtPassword") <> "" then
                Session("STxtuseremail") = rslogin("UserEmail")
                Session("STxtUserName") = rslogin("UserFullName")
                Session("STxtPassword")= RSlogin("Password")
                Session("SIsAdmin") = RSLogin("IsAdmin")
                Session("SLoggedDT") = Now()

                'Application Object Start
                    'if Application("ActiveUsers") = "" then
                        'Application("ActiveUsers") = 1
                    'else
                        'Application("ActiveUsers") = Application("ActiveUsers") + 1
                    'end if 
                'Application Object end
                

                'Creating Cookie
                if request.form("RememberMe")="1" then
                    response.cookies("CookieUserName")=RSLogin("UserEmail")
                    response.cookies("CookieUserName").Expires=#7-Jan-2021#
                else
                    response.cookies("CookieUserName")=""
                    response.cookies("CookieUserName").Expires=Now()
                end if 
                
                response.Redirect("Menu.asp")
            end if
        end if
    end if
%>

