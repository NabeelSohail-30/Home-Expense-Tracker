<%
    'Start - Validate logged user Session Timeout. If Session is timeout then it will redirect to te login page.
        Session("STimeoutError")=""

        if Session("STxtUserEmail")="" then
            Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
            response.Redirect("Login.asp")
        end if
    'end
%>
    <!--Start - Validate the Logged User Access Everytime on any page access.-->
    <!--Use of Include File method will save time to rewrite the code on every page.-->
        <!--#include file="ValidateLogin.asp"-->
    <!--End-->
<%
    'Variables
        Dim FromAcc
        Dim ToAcc
        Dim TransferDate
        Dim TransferDesc
        Dim TransferAmount

        Dim RSAccountFrom
        Dim RSAccountTo

        Dim ErrorFound
    'End

    'Initializing
        FromAcc = Request.Form("FormAccountFrom")
        ToAcc = Request.Form("FormAccountTo")
        TransferDate = Request.Form("TrDate")
        TransferDesc = Request.Form("TrDesc")
        TransferAmount = Request.Form("TransAmount")
    'End

    'response.Write(FromAcc)
    'response.Write(ToAcc)
    'response.Write(TransferDate)
    'response.Write(TransferDesc)
    'response.Write(TransferAmount)

    Session("ErrorFound")=""
    Session("DateError")=

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'TableName Validation Start
        If FromAcc = "" or ToAcc = "" Then
            Session("ErrorTable")="Invalid Table Name! Blank Table Name Found"
            ErrorFound=True
        Else
            Session("ErrorFound")=""
        End If
    'TableName Validation End
    
    'Date Validation Start
        if TransferDate = "" then
            Session("DateError")="Enter Transfer Date"
            ErrorFound=true
        end if

        if TransDate <> "" then
            If IsDate(TransferDate)=False Then
                Session("DateError")="Input is not Date"
                Call CloseConn()
                response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
            Else
                Session("DateError")=""
            End If
        End If
    'Date Validation Ends
    
    'Currency Validation Starts
            If TransferAmount = "" Then
                Session("CreditError")="You Cannot Leave Transfer Amount NULL"
                ErrorFound=true
            ElseIf IsNumeric(TransferAmount) = False Then
                Session("CreditError")="Character Found in Transfer Amount"
                ErrorFound=true
            ElseIf TransferAmount <= 0 Then
                Session("CreditError")="Transfer Amount cannot be less than or equal to zero"
                ErrorFound=true
            End If
    'Currency Validation Ends

    if ErrorFound=true then
        Call CloseConn()
        response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
    end if

    'Inserting Rec in FromAccount
        Set RSAccountFrom = Server.CreateObject("ADODB.RecordSet")
        RSAccountFrom.Open "SELECT  Top (1) Balance FROM " & FromAcc & " ORDER BY ID DESC",Conn

        Dim LastBal
        Dim CurBal

        LastBal = RSAccountFrom("Balance")
        'response.Write("LastBal = " & Lastbal)
        'response.End

        CurBal = (LastBal - TransferAmount) + 0

        QryStr = "INSERT INTO " & FromAcc & " (TransactionDate,CategoryID,PersonID,Description,Credit,Debit,Balance) Values ('" & TransferDate & "', 18 , 1" & ",'" & TransferDesc & "'," & _
                  TransferAmount & ", 0 ," & CurBal & ")"

        'response.Write qrystr
        'response.end
        
        Conn.Execute QryStr
    'End

    'Inserting Rec in FromAccount
        Set RSAccountTo = Server.CreateObject("ADODB.RecordSet")
        RSAccountTo.Open "SELECT  Top (1) Balance FROM " & ToAcc & " ORDER BY ID DESC",Conn

        'Dim LastBal
        'Dim CurBal

        LastBal = RSAccountTo("Balance")
        'response.Write("LastBal = " & Lastbal)
        'response.End

        CurBal = (LastBal - 0) + TransferAmount

        QryStr = "INSERT INTO " & ToAcc & " (TransactionDate,CategoryID,PersonID,Description,Credit,Debit,Balance) Values ('" & TransferDate & "', 18 , 1" & ",'" & TransferDesc & "', 0 , " & TransferAmount & ", " & CurBal & ")"

        'response.Write qrystr
        'response.end
        
        Conn.Execute QryStr
    'End

    'Closing RS
        RSAccountTo.Close
        Set RSAccountTo = Nothing

        RSAccountFrom.Close
        Set RSAccountFrom = Nothing
    'End

    response.Redirect("Transfer.asp")

%>