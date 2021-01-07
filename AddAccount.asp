<%
    '1. Session Timeout
    '2. Storing Form Values in Variables
    '3. Validations
    '4. INSERT INTO ACCOUNTSTABLE
    '5. CREATE NEW ACC TABLE
    '6. INSERT opening bal data in new account table

    Session("STimeoutError")=""

    if Session("STxtUserEmail")="" then
        Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
        response.Redirect("Login.asp")
    end if
%>
    <!--#include file="ValidateLogin.asp"-->
<%

    'Variable Declaration Start
        Dim AccountName
        Dim AccDesc
        'Dim Conn
        Dim OpnDate
        Dim OpnBal
        'Dim QryStr
    'Variable Declaration End

    'Variable Initialization Start
        AccountName = request.Form("AccName")
        AccDesc = request.form("AccDescription")
        OpnDate = request.Form("OpnBalDate")
        OpnBal = request.Form("OpnBalance")
    'Variable Initialization End

    'Validations
    If AccountName = "" then
        Session("ErrorAcc")="Account Name cannot be Null"
    else
        Session("ErrorAcc")=""
    end if

    if InStr(AccountName," ") > 0 then
        Session("ErrorAcc")="Account Name cannot contain space"
    else
        Session("ErrorAcc")=""
    end if

    If AccDesc = "" then
        Session("ErrorAcc")="Account Description cannot be Null"
    else
        Session("ErrorAcc")=""
    end if

    If OpnDate = "" then
        Session("ErrorDate")="Opening Date cannot be Null"
    else
        Session("ErrorDate")=""
    end if

    if OpnDate <> "" then
        if isdate(OpnDate)=false then
            Session("ErrorDate")="No Date Found in Opening Date"
        else
            Session("ErrorDate")=""
        end if
    end if

    If OpnBal = "" then
        Session("ErrorBal")="Opening Balance cannot be Null"
    else
        Session("ErrorBal")=""
    end if

    If OpnBal <> "" then
        if isnumeric(OpnBal)=False then
            Session("ErrorBal")="Invalid Balance Found!"
        elseif OpnBal < 0 then
            Session("ErrorBal")="Balance Cannot be less than zero"
        end if
    end if

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'Inserting New Table Start
        QryStr = "INSERT INTO AccountsTable (TableName, TableDescription) Values('" & AccountName & "','" & AccDesc & "')"
        'response.Write qrystr
        Conn.Execute QryStr
        'response.Write("New Account Added")
    'Inserting New Table End
    
    'Creating Table Start
        QryStr = "CREATE TABLE " & Accountname & "(" & _
                  "ID int IDENTITY(1,1) NOT NULL," & _
                  "TransactionDate datetime NOT NULL," & _
                  "CategoryID int NOT NULL," & _
                  "PersonID int NOT NULL," & _
                  "Description varchar(100) NOT NULL," & _
                  "Credit money NOT NULL," & _
                  "Debit money NOT NULL," & _
                  "Balance money NOT NULL," & _
                  "CreationDateTime datetime NOT NULL CONSTRAINT DF_" & AccountName & "_CreationDateTime DEFAULT (getdate())," & _
                  "CONSTRAINT PK_" & AccountName & " PRIMARY KEY CLUSTERED" & _
                  "(ID ASC) " & _
                  "WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"         
    
        'Response.Write("<br>" & QryStr)
    
        Conn.Execute QryStr
        'response.Write("New Account Created")
    'Creating Table End

    'Inserting Rec Start
        QryStr = "INSERT INTO " & AccountName & " (TransactionDate,CategoryID,PersonID,Description,Credit,Debit,Balance) Values ('" & OpnDate & "'," & 14 & "," & 1 & ",'Opening Balance'," & _
                  0 & "," & OpnBal & "," & OpnBal & ")"
        
        'response.write("<br>" & qrystr)
        Conn.Execute QryStr
        'response.Write("Opening Rec Added")
    'Inserting Rec End   

    response.Redirect("Menu.asp?Msg=New Account (" & AccountName & ") has been added successfully")
%>