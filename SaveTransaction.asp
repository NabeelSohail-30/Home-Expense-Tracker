<%
    '1. Store all the form values (request.form) into memory variables
    '2. Validations (same asa ms-access)
    '3. Open DB
    '4. Open RS
    '5. Calculate Balance 
        '5.1 Get Last Balance through RS
        '5.2 Calculate RS Last Balance with Form credit or debit values
    '5. Build INSERT qry 
    '6. Execute Insert Qry through Conn.Execute method
    '7. Redirect to NewTransaction page with Success or Error Msg
    '8. Incase of error all previous form value should be restored

    'response.Write("Hello")

    'response.end

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
    Dim TransDate
    Dim Catg
    Dim Person
    Dim Description
    Dim CrAmount
    Dim DbAmount

    'Dim ErrorFound
    Dim TableFound
    Dim CatgFound
    Dim PersonFound

    'Dim Conn 
    'Dim CS
    Dim RSTable
    Dim RSPerson
    Dim RSCatg
    Dim RSTransaction
    'Variable Declaration End

    'Session Variables Start
    Session("DateError")=""
    Session("CreditError")=""
    Session("DebitError")=""
    Session("ErrorFound")=""
    'Session Variables End

    'Initializing Variables Start
    TableName = request.Form("AccTableName")
    TransDate = request.Form("TrDate")
    Catg = request.Form("SelectCategory")
    Person = request.Form("SelectPerson")
    Description = request.Form("TrDesc")
    CrAmount = request.Form("CreditAmount")
    DbAmount = request.Form("DebitAmount")
    ErrorFound=false
    TableFound=False
    CatgFound=False
    PersonFound=False
    'Initializing Variables End

    'response.Write(TableName)
    'response.Write(TransDate)
    'response.Write(Catg)
    'response.Write(Person)
    'response.Write(Description)
    'response.Write(CrAmount)
    'response.Write(DbAmount)
    'response.Write(ErrorFound)
    'response.Write(TableFound)
    'response.Write(CatgFound)
    'response.Write(PersonFound)
    
    'response.end

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'TableName Validation Start
        If TableName = "" Then
            Session("ErrorFound")="Invalid Table Name! Blank Table Name Found"
            response.Redirect("NewTransaction.asp")
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
            response.Redirect("NewTransaction.asp")
        End If
    'TableName Validation End
    
    'Catg Validation Start
        if Catg = "" Then
            Session("ErrorFound")="Catg cannot be left NULL"
            response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If
        
        Set RSCatg = Server.CreateObject("ADODB.RecordSet")
        rscatg.Open "SELECT CategoryID From Categories",Conn
        
        
        Do While RSCatg.EOF=False
            'response.write("<br>" & catg & "-" & rscatg("categoryid"))
            'response.Write("<br>" & cint(Catg) = RSCatg("CategoryID"))
            IF Cint(Catg) = RSCatg("CategoryID") Then
                CatgFound=True
                exit do
            End If
            RSCatg.MoveNext
        Loop
        'response.Write("end of loop")
        'response.End
        If CatgFound=False then
            Session("ErrorFound")="Category not Found"
            Call CloseRSCatg()
            Call CloseConn()
            response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If
    'Catg Validation End

    'Person Validation Start
        if Person = "" Then
            Session("ErrorFound")="Person cannot be left NULL"
            response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
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
            response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
        else
            Session("ErrorFound")=""
        End If
    'Person Validation Ends
    'response.Write("Table, person and Catg Validation ok")
    'response.End
    'Date Validation Start
        if TransDate = "" then
            Session("DateError")="Enter Transaction Date"
            ErrorFound=true
        end if

        if TransDate <> "" then
            If IsDate(TransDate)=False Then
                Session("DateError")="Input is not Date"
                Call CloseConn()
                response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
            Else
                Session("DateError")=""
            End If
        End If

        
    'Date Validation Ends
    
    'Currency Validation Starts
        If (CrAmount <> "") And (DbAmount <> "") Then
            Session("CreditError")="You Cannot Provide Credit and Debit Together"
            ErrorFound=true
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
            ElseIf CrAmount <= 0 Then
                Session("CreditError")="Credit cannot be less than or equal to zero"
                ErrorFound=true
            End If
            DbAmount = 0
        End If
    
        If CrAmount = "" Then
            If DbAmount = "" Then
                Session("DebitError")="You Cannot Leave Debit NULL"
                ErrorFound=true
            ElseIf IsNumeric(DbAmount) = False Then
                Session("DebitError")="Character Found in Debit"
                ErrorFound=true
            ElseIf DbAmount <= 0 Then
                Session("DebitError")="Debit cannot be less than or equal to zero"
                ErrorFound=true
            End If
        CrAmount = 0
        End If
    'Currency Validation Ends

    if ErrorFound=true then
        Call CloseConn()
        response.Redirect("NewTransaction.asp?AccTableName=" & TableName)
    end if
    
    'Calculating Bal Start
        Set RSTransaction = Server.CreateObject("ADODB.RecordSet")
        RSTransaction.Open "SELECT  Top (1) Balance FROM " & TableName & " ORDER BY ID DESC",Conn

        Dim LastBal
        Dim CurBal

        LastBal = RSTransaction("Balance")
        'response.Write("LastBal = " & Lastbal)
        'response.End

        CurBal = (LastBal - CrAmount) + DbAmount
    'Calculating Bal End

    'Inserting Rec Start
        'Dim QryStr

        QryStr = "INSERT INTO " & TableName & " (TransactionDate,CategoryID,PersonID,Description,Credit,Debit,Balance) Values ('" & TransDate & "'," & Catg & "," & Person & ",'" & Description & "'," & _
                  CrAmount & "," & DbAmount & "," & CurBal & ")"

        'response.Write qrystr
        'response.end
        
        Conn.Execute QryStr
    'Inserting Rec End
        
    'Closing all RS
        
    'End of RS Closing
        Call CloseRSTransaction()
        Call CloseRSTable()
        Call CloseRSCatg()
        Call CloseRSPerson()
    'Conn Closing
        Call CloseConn()

    'Redirecting to View Transaction
        Response.Redirect("ViewTransaction.asp?AccTableName=" & TableName)

    'Procedure to close db Conn
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

%>