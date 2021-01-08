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

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'Inserting Rec in FromAccount
        Set RSTransaction = Server.CreateObject("ADODB.RecordSet")
        RSTransaction.Open "SELECT  Top (1) Balance FROM " & FromAcc & " ORDER BY ID DESC",Conn

        Dim LastBal
        Dim CurBal

        LastBal = RSTransaction("Balance")
        'response.Write("LastBal = " & Lastbal)
        'response.End

        CurBal = (LastBal - TransferAmount) + 0

        QryStr = "INSERT INTO " & FromAcc & " (TransactionDate,CategoryID,PersonID,Description,Credit,Debit,Balance) Values ('" & TransferDate & "', 18 , 1" & ",'" & TransferDesc & "'," & _
                  TransferAmount & ", 0 ," & CurBal & ")"

        response.Write qrystr
        'response.end
        
        'Conn.Execute QryStr
    'End

%>