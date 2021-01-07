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
    Dim RSTransaction

    Set Conn = Server.CreateObject("ADODB.Connection")
    Set RSTransaction = Server.CreateObject("ADODB.RecordSet")

    CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
    Conn.Open CS
   
    Dim TransID
    Dim TableName
    Dim TransDate
    Dim TransDateYYMMDD
    Dim Catg
    Dim Person
    Dim TransDesc
    Dim CrAmount
    Dim DbAmount
    'Dim QryStr
    Dim ActionNumber

    TransID = Request.QueryString("QsID")
    TableName = request.QueryString("QSTableName")
    ActionNumber = request.QueryString("Action")
    
    If TransId = "" Then
        Session("ErrorQsID")="Invalid Transaction ID! Blank Transaction Id Found "
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    Elseif IsNumeric(TransId) = False then
        Session("ErrorQsID")="Invalid Transaction ID! Character Found in Transaction Id"
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    ElseIf Cdbl(TransID) < 1 Then
        Session("ErrorQsID")="Invalid Transaction ID! Transaction ID cannot be less than 1"
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    Else
        Session("ErrorQsID")=""
    End If
    
    If TableName = "" Then
        Session("ErrorTableName")="Invalid Table Name! Blank Table Name Found"
        response.Redirect("ViewTransaction.asp")
    Else
        Session("ErrorTableName")=""
    End If

    Dim TableFound
    Dim RSTable
    Set RSTable = Server.CreateObject("ADODB.RecordSet")
    TableFound=False

    If RSTable.State=1 Then
        RSTable.Close
    End If
    RSTable.Open "SELECT TableName FROM AccountsTable",Conn

    Do While RSTable.EOF=False
        IF LCase(TableName) = LCase(RSTable("TableName")) Then
            TableFound=True
        End If
        RSTable.MoveNext
    Loop

    If TableFound=False Then
        Session("ErrorTableName")="Invalid Table Name! Table Name Not Found"
        Response.Redirect("ViewTransaction.asp")
    End If

    If ActionNumber = "" Then
        Session("ErrorAction")="Invalid Action"
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    Elseif IsNumeric(ActionNumber)=False then
        Session("ErrorAction")="Invalid Action"
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    Elseif ActionNumber <> 1 AND ActionNumber <> 2 then
        Session("ErrorAction")="Invalid Action"
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    Else
        Session("ErrorAction")=""
    End If
    
    'response.Write(TransID)
    'response.Write(TableName)
     
    QryStr = "SELECT * FROM " & TableName & " WHERE (ID = " & TransID & ")"

    'response.Write(qrystr)

    RSTransaction.Open QryStr,Conn

    Dim Date_Day
    Dim Date_Mon
    Dim Date_Year

    If RSTransaction.EOF Then
        Session("ErrorQsId")="Transaction Id Not Found"
        Response.Redirect("ViewTransaction.asp?AccTableName="&TableName)
    Else
        Session("ErrorQsID")=""
    End If
        
    TransDate = RSTransaction("TransactionDate")
    Date_Day = Day(TransDate)
    Date_Mon = Month(TransDate)
    Date_Year = Year(TransDate)

    If Date_Day >=1 AND Date_Day <=9 Then
        Date_Day = "0" & Date_Day
    End If

    If Date_Mon >=1 AND Date_Mon <=9 Then
        Date_Mon = "0" & Date_Mon
    End If
    'response.Write("day = " & Date_Day)
    'response.Write("mon = " & Date_Mon)
    'response.Write("year = " & Date_Year)
    TransDateYYMMDD = Date_Year & "-" & Date_Mon & "-" & Date_Day
    'Response.Write(TransDateYYMMDD)

    Catg = RSTransaction("CategoryID")
    Person = RSTransaction("PersonID")
    TransDesc = RSTransaction("Description")
    CrAmount = RSTransaction("Credit")
    DbAmount = RSTransaction("Debit")

    'response.Write("Table = " & TableName)
    'response.Write("ID = " & TransID)
    'response.Write("Date = " & TransDate)
    'response.Write("Catg = " & Catg)
    'response.Write("Person = " & Person)
    'response.Write("Desc = " & TransDesc)
    'response.Write("Credit = " & CrAmount)
    'response.Write("Debit = " & DbAmount)

%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <% If ActionNumber = 1 Then %>
    <title>Home Expense Tracker - Edit Transaction</title>
    <% Else %>
    <title>Home Expense Tracker - Delete Transaction</title>
    <% End If %>
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <link rel="stylesheet" href="CSS/StyleNewTrans.css">
</head>

<body>
    <!--- #include file="Header.asp" -->

    <div class="container-fluid">
        <div class="row">
            <div class="col">
                <div class="wrapper">
                    <header>
                        <% If ActionNumber = 1 Then %>
                        <h1>Edit Transaction</h1>
                        <% Else %>
                        <h1>Delete Transaction</h1>
                        <% End If %>
                    </header>

                    <hr>

                    <% If ActionNumber = 1 Then %>
                    <form action="UpdateRec.asp" method="POST" class="combo-account">
                        <% Else %>
                        <form action="DeleteRec.asp" method="POST" class="combo-account">
                            <% End If %>
                            <input type="hidden" name="TransID" value="<% Response.write(TransID) %>" />
                            <input type="hidden" name="AccTableName" value="<% Response.write(TableName) %>" />
                            <div>
                                <label for="AccountTableName">Select Account</label>
                                <select name="AccountTableName" class="select-account form-control" disabled>
                                    <%
                                    Dim RSAccounts
                                    Set RSAccounts = Server.CreateObject("ADODB.RecordSet")
                
                                    RSAccounts.Open "SELECT TableName, TableDescription FROM AccountsTable",Conn
        
                                    do while Not RSAccounts.EOF
        
                                    If RSAccounts("TableName") = TableName then
                               %>
                                    <option value="<% response.write(RSAccounts("TableName")) %>" selected>
                                        <% response.write(RSAccounts("TableDescription")) %></option>
                                    <%
                                    else  
                               %>
                                    <option value="<% response.write(RSAccounts("TableName")) %>">
                                        <% response.write(RSAccounts("TableDescription")) %></option>
                                    <%
        
                                   End if
        
                                   RSAccounts.MoveNext
                                   Loop
        
                                   RSAccounts.Close
                                   Set RSAccounts = Nothing
                              %>
                                </select>
                            </div>

                            <hr>

                            <div class="form-group">
                                <label for="">Transaction Date</label>
                                <input type="date" class="form-control" required name="TrDate"
                                    value="<% Response.write(TransDateYYMMDD) %>">
                            </div>

                            <div class="form-group">
                                <label for="">Category</label>
                                <select name="SelectCategory" class="select-account form-control">
                                    <%
                                Dim RSCatg
                                Set RSCatg = Server.CreateObject("ADODB.RecordSet")
                
                                RSCatg.Open "SELECT CategoryID, Category FROM Categories",Conn
        
                                do while Not RSCatg.EOF 
        
                                IF RSCatg("CategoryID") = Catg then
                            %>
                                    <option value="<% response.write(RSCatg("CategoryID")) %>" selected>
                                        <% response.write(RSCatg("Category")) %></option>
                                    <% Else %>
                                    <option value="<% response.write(RSCatg("CategoryID")) %>">
                                        <% response.write(RSCatg("Category")) %></option>
                                    <%
                                End if
        
                                RSCatg.MoveNext
                                Loop
        
                                RSCatg.Close
                                Set RSCatg = Nothing
                            %>
                                </select>
                            </div>

                            <div class="form-group">
                                <label for="">Person</label>
                                <select name="SelectPerson" class="select-account form-control">
                                    <%
                                Dim RSPerson
                                Set RSPerson = Server.CreateObject("ADODB.RecordSet")
                
                                RSPerson.Open "SELECT PersonID, PersonName FROM Persons",Conn
        
                                do while Not RSPerson.EOF 
        
                                If RSPerson("PersonID") = Person then
                            %>
                                    <option value="<% response.write(RSPerson("PersonID")) %>" selected>
                                        <% response.write(RSPerson("PersonName")) %></option>
                                    <% Else %>
                                    <option value="<% response.write(RSPerson("PersonID")) %>">
                                        <% response.write(RSPerson("PersonName")) %></option>
                                    <%
                                End if
                                RSPerson.MoveNext
                                Loop
        
                                RSPerson.Close
                                Set RSPerson = Nothing
                            %>
                                </select>
                            </div>

                            <div class="form-group">
                                <label for="">Transaction Description</label>
                                <textarea cols="30" rows="3" class="form-control"
                                    name="TrDesc"><% Response.Write(TransDesc) %></textarea>
                            </div>

                            <div class="form-group">
                                <label for="">Credit Amount</label>
                                <input type="text" class="form-control" name="CreditAmount"
                                    value="<% Response.write(CrAmount) %>">
                            </div>

                            <div class="form-group">
                                <label for="">Debit Amount</label>
                                <input type="text" class="form-control" name="DebitAmount"
                                    value="<% Response.write(DbAmount) %>">
                            </div>

                            <div>
                                <% If ActionNumber = 1 Then %>
                                <input type="submit" value="Edit Record" class="btn btn-primary">
                                <% Else %>
                                <input type="submit" value="Delete Record" class="btn btn-danger ml-2">
                                <% End If %>
                            </div>
                        </form>
                </div>
            </div>
        </div>
    </div>

<!--- #include file="footer.asp" -->
</body>

</html