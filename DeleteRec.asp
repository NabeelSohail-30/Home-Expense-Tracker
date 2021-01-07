<%
Session("STimeoutError")=""

if Session("STxtUserEmail")="" then
    Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
    response.Redirect("Login.asp")
end if
%>
<!--#include file="ValidateLogin.asp"-->
<%
    Dim TableName
    Dim TransID
    Dim TransDate
    Dim Catg
    Dim Person
    Dim Description
    Dim CrAmount
    Dim DbAmount
    'Dim ErrorFound

    TableName = request.Form("AccTableName")
    TransID = Request.Form("TransID")
    'TransDate = request.Form("TrDate")
    'Catg = request.Form("SelectCategory")
    'Person = request.Form("SelectPerson")
    'Description = request.Form("TrDesc")
    'CrAmount = request.Form("CreditAmount")
    'DbAmount = request.Form("DebitAmount")

    'Req Validations
    'TableName <> null
    'TableName not found
    'Transid <> null
    'id not found

    'Dim Conn 
    'Dim CS
    Dim RSTransaction

    Set Conn = Server.CreateObject("ADODB.Connection")
    Set RSTransaction = Server.CreateObject("ADODB.RecordSet")

    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS

    'Dim QryStr

    QryStr = "DELETE FROM " & TableName & " WHERE(ID = " & TransID & ")"
    'response.Write(qrystr)
    Conn.Execute QryStr

    QryStr = "SELECT Top (1) Balance FROM " & TableName & " WHERE(ID < " & TransID & ") ORDER BY ID DESC"
    'response.Write("<br>" & qrystr)
    'response.End

    If RSTransaction.State = 1 Then
        RSTransaction.close
    End If
    RStransaction.Open QryStr,Conn

    Dim BalAmount

    If RStransaction.BOF Or RStransaction.EOF Then
        BalAmount = 0
    Else
        BalAmount = RStransaction("Balance")
    End If

    'Calculating Balance for rest of the Records
    QryStr = "SELECT ID,Credit,Debit FROM " & TableName & " WHERE (ID > " & TransID & ") ORDER BY ID ASC"
    'response.Write("<br>" & qrystr)

    If RSTransaction.State = 1 Then
        RSTransaction.close
    End If
    RStransaction.Open QryStr,Conn

    Dim CreditAmount
    Dim DebitAmount
    Dim TrID
    
    Do While RStransaction.EOF = False       
        TrID = RStransaction("ID")
        CreditAmount = RStransaction("Credit")
        DebitAmount = RStransaction("Debit")
    
        BalAmount = (BalAmount - CreditAmount) + DebitAmount

        qrystr = "Update " & TableName & " SET Balance = " & BalAmount & " WHERE (ID = " & TrID & ")"
        'response.Write("<br>" & qrystr)
        Conn.Execute QryStr

        RStransaction.MoveNext
    Loop
    
    Response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
%>