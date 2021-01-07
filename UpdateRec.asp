<%
Session("STimeoutError")=""

if Session("STxtUserEmail")="" then
    Session("STimeoutError")="Your Session has been Timed Out! Please Login to continue"
    response.Redirect("Login.asp")
end if
%>
<!--#include file="ValidateLogin.asp"-->
<%

    'Variable Declaration Start
        Dim TableName
        Dim TransID
        Dim TransDate
        Dim Catg
        Dim Person
        Dim Description
        Dim CrAmount
        Dim DbAmount

        'Dim ErrorFound
        Dim TableFound
        Dim IdFound
        Dim CatgFound
        Dim PersonFound

        'Dim Conn 
        'Dim CS
        Dim RSTrId
        Dim RSTable
        Dim RSPerson
        Dim RSCatg
        Dim RSTransaction
    'Variable Declaration End

    'Session Variable Start
        Session("ErrorFound")=""
        Session("DateError")=""
        Session("CreditError")=""
        Session("DebitError")=""
    'Session Variable End

    'Variable Initialization Start
        TableName = request.Form("AccTableName")
        TransID = Request.Form("TransID")
        TransDate = request.Form("TrDate")
        Catg = request.Form("SelectCategory")
        Person = request.Form("SelectPerson")
        Description = request.Form("TrDesc")
        CrAmount = request.Form("CreditAmount")
        DbAmount = request.Form("DebitAmount")
        TableFound=False
        ErrorFound=False
        CatgFound=False
        PersonFound=False
        IdFound = False
    'Variable Initialization End

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'TableName Validation Start
        If TableName = "" Then
            Session("ErrorFound")="Invalid Table Name! Blank Table Name Found"
            response.Redirect("ViewTransaction.asp")
        Else
            Session("ErrorFound")=""
        End If

        Set RSTable = Server.CreateObject("ADODB.RecordSet")
        RSTable.Open "SELECT TableName FROM AccountsTable",Conn

        Do While RSTable.EOF=False
            IF LCase(TableName) = LCase(RSTable("TableName")) Then
                TableFound=True
                Exit Do
            End If
            RSTable.MoveNext
        Loop

        If TableFound=False Then
            Session("ErrorFound")="Invalid Table Name! Table Name Not Found"
            Call CloseRSTable()
            Call CloseConn()
            response.Redirect("ViewTransaction.asp")
        End If
    'TableName Validation End
    
    'Transaction ID Start
        If TransID = "" then
            Session("ErrorFound")="Invalid Transaction Id! Blank ID Found"
            Response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
        Else
            Session("ErrorFound")=""
        End If

        Set RSTrId = Server.CreateObject("ADODB.RecordSet")
        RSTrId.Open "SELECT Id FROM " & TableName,Conn

        Do While RSTrId.EOF=False
            IF cint(TransId) = RSTrId("Id") Then
                IdFound=True
                Exit Do
            End If
            RSTrId.MoveNext
        Loop

        If IdFound=False Then
            Session("ErrorFound")="Invalid Transaction Id! ID not Found"
            Call CloseRSTrId
            Call CloseConn()
            Response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
        End If
    'Transaction ID End

    'Catg Validation Start
        if Catg = "" Then
            Session("ErrorFound")="Catg cannot be left NULL"
            response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If
        
        Set RSCatg = Server.CreateObject("ADODB.RecordSet")
        rscatg.Open "SELECT CategoryID From Categories",Conn
        
        Do While RSCatg.EOF=False
            IF Cint(Catg) = RSCatg("CategoryID") Then
                CatgFound=True
                exit do
            End If
            RSCatg.MoveNext
        Loop
        
        If CatgFound=False then
            Session("ErrorFound")="Category not Found"
            Call CloseRSCatg()
            Call CloseConn()
            response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If
    'Catg Validation End

    'Person Validation Start
        if Person = "" Then
            Session("ErrorFound")="Person cannot be left NULL"
            response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If

        Set RSPerson = Server.CreateObject("ADODB.RecordSet")
    
        RSPerson.Open "SELECT PersonID From Persons",Conn

        Do While RSPerson.EOF=False
            IF cint(Person) = RSPerson("PersonId") Then
                PersonFound=True
                exit do
            End If
            RSPerson.MoveNext
        Loop

        If PersonFound=False then
            Session("ErrorFound")="Person not Found"
            Call CloseRSPerson()
            Call CloseConn()
            response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If
    'Person Validation Ends
    
    'Date Validation Start
        if TransDate = "" then
            Session("DateError")="Enter Transaction Date"
            ErrorFound=true
        end if

        if TransDate <> "" then
            If IsDate(TransDate)=False Then
                Session("DateError")="Input is not Date"
                Call CloseConn()
                response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
            Else
                Session("DateError")=""
            End If
        End If
    'Date Validation Ends
    
    'Currency Validation Starts
            

        If (CrAmount <> "") And (DbAmount <> "") Then
            if (CrAmount <> 0) AND (DbAmount <> 0) then
                Session("CreditError")="You Cannot Provide Credit and Debit Together"
                ErrorFound=true
            end if
        End If
    
        If (CrAmount = "") And (DbAmount = "") Then
            Session("CreditError")="You Cannot Leave Credit and Debit null together"
            ErrorFound=true
        End If
    
        If DbAmount = "" Then
            If CrAmount = "" Then
                Session("CreditError")="You Cannot Leave Credit NULL"
                ErrorFound=true
            ElseIf IsNumeric(CrAmount) = False Then
                Session("CreditError")="Character Found in Credit"
                ErrorFound=true
            ElseIf CrAmount < 0 Then
                Session("CreditError")="Credit cannot be less than or equal to zero"
                ErrorFound=true
            End If
            DbAmount=0
        End If
    
        If CrAmount = "" Then
            If DbAmount = "" Then
                Session("DebitError")="You Cannot Leave Debit NULL"
                ErrorFound=true
            ElseIf IsNumeric(DbAmount) = False Then
                Session("DebitError")="Character Found in Debit"
                ErrorFound=true
            ElseIf DbAmount < 0 Then
                Session("DebitError")="Debit cannot be less than or equal to zero"
                ErrorFound=true
            End If
            CrAmount=0
        End If
    'Currency Validation Ends

    if ErrorFound=true then
        Call CloseConn()
        response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)
    end if


    'Updating Record Start
        Set RSTransaction = Server.CreateObject("ADODB.RecordSet")

        'Dim QryStr

        'response.Write("Table Name = " & request.Form("AccTableName"))
        QryStr = "SELECT Top (1) Balance FROM " & TableName & " WHERE(ID < " & TransID & ") ORDER BY ID DESC"
        'response.Write("<br>" & qrystr)
        'response.End

        'Calculating Balance
        RSTransaction.Open QryStr,Conn

        Dim LastBal
        Dim CurrBal

        If RStransaction.BOF Or RStransaction.EOF Then
            LastBal = 0
        Else
            LastBal = RStransaction("Balance")
        End If 

        CurrBal = (LastBal - CrAmount) + DbAmount

        QryStr = "Update " & TableName & " SET TransactionDate = '" & TransDate & "', CategoryID = " & Catg & _
                ", PersonID = " & Person & ", Description = '" & Description & "', Credit = " & CrAmount & _
                ", Debit = " & DbAmount & ",Balance = " & CurrBal & " WHERE (ID = " & TransID & ")" 

        'response.write("<br>" & QryStr)

        Conn.Execute QryStr

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
    
            CurrBal = (CurrBal - CreditAmount) + DebitAmount

            qrystr = "Update " & TableName & " SET Balance = " & CurrBal & " WHERE (ID = " & TrID & ")"
            'response.Write("<br>" & qrystr)
            Conn.Execute QryStr

            RStransaction.MoveNext
        Loop
    'Updating Record Ends

    'RS Closing
        Call CloseRSTransaction()
        Call CloseRSTable()
        Call CloseRSCatg()
        Call CloseRSPerson()
        Call CloseRSTrId()
    'Conn Closing
        Call CloseConn()
    
    'Dim BuildUrl
    'BuildUrl = "ViewTransaction.asp?AccTableName=" & TableName
    
    'Need to discuss 1 big problem for edit page
    Response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)

    Sub CloseConn()
        Conn.close
        Set Conn = Nothing
    End Sub

    Sub CloseRSTransaction()
        if RSTransaction.State=1 then
            RSTransaction.Close
        end if
        Set RSTransaction = Nothing
    End Sub

    Sub CloseRSTable()
        if RSTable.State=1 then
            RSTable.Close
        end if
        Set RsTable = Nothing
    End Sub

    Sub CloseRSCatg()
        if RSCatg.State=1 then
            RSCatg.Close
        end if
        Set RSCatg = Nothing
    End Sub

    Sub CloseRSPerson()
        if RSPerson.State=1 then
            RSPerson.Close
        end if
        Set RSPerson = Nothing
    End Sub

    Sub CloseRSTrId()
        if RSTrId.State=1 then
            RSTrId.Close
        end if
        Set RSTrId = Nothing
    End Sub
%>