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
    'Dim QryStr
    Dim RecPerPage
    Dim RecNumber
    Dim PageNumber
    Dim SkipRec
    Dim LastPage

    'Opening Db Start
        Set Conn = Server.CreateObject("ADODB.Connection")
        CS = "Driver={SQL Server};Server=NABEELS-WORK;Database=HomeExpenseTracker;User Id=homeexpense;Password=Nabeel30;"
        Conn.Open CS
    'Opening Db End

    'Del Pers Start
        If request.QueryString("Action")= "2" AND Request.QueryString("QsId") <> "" then
            QryStr = "DELETE FROM Persons WHERE (PersonId = " & cint(Request.querystring("QsId")) & ")" 
            'response.Write qrystr
            Conn.Execute QryStr
        end if
    'Del Pers End

    'Save new Pers Start
        If request.Form("TxtPersName") <> "" And Request.Form("TxtPersId") = "" then
            'response.Write("Form Posted with value")
            
            QryStr="INSERT INTO Persons (PersonName) Values('" & Request.Form("TxtPersName") & "')"
            'response.Write qrystr
            
            Conn.Execute QryStr
        end if
    'Save New Pers End

    'Edit Pers Start
        If request.Form("TxtPersName") <> "" AND Request.form("TxtPersId") <> "" then
            'response.Write("Form Posted with value")
            
            QryStr="UPDATE Persons SET PersonName = '" & Request.Form("TxtPersName") & "' WHERE (PersonId = " & cint(Request.form("TxtPersId")) & ")"
            'response.Write qrystr
            
            Conn.Execute QryStr
        end if
    'Edit Pers End

    'Get Value for Edit
        If request.QueryString("Action")= "1" AND Request.QueryString("QsId") <> "" then
            Dim RSEditPers
            Set RSEditPers = Server.CreateObject("ADODB.RecordSet")
            
            QryStr = "SELECT * FROM Persons WHERE (PersonId = " & cint(Request.QueryString("QsId")) & ")"
            'response.Write qrystr

            RSEditPers.Open QryStr,Conn
      
            Dim PersonName
            Dim PersonId
          
            PersonName = RSEditPers("PersonName")
            PersonId = RSEditPers("PersonId")  
    
            'response.write PersonName  

            RSEditPers.Close
            Set RSEditPers = Nothing
        end if

    

    Dim RSCount
    Set RSCount = Server.CreateObject("ADODB.RecordSet")
    
    RSCount.Open "SELECT COUNT(PersonId) AS TotalRecords FROM Persons",Conn

    TotalRec=RSCount("TotalRecords")
    RecPerPage=5

    If RSCount.EOF  or RSCount("TotalRecords")=1 then
        LastPage = 0
    else
        LastPage = Cstr((RSCount("TotalRecords")/RecPerPage))

        If InStr(LastPage,".") > 1 then
            LastPage = cint(LEFT(LastPage,InStr(LastPage,".")-1)) + 1
        end if
    End If

    If Request.QueryString("QsPageNumber")="" then
        PageNumber = 1
        SkipRec=0
    else
        PageNumber = Cint(request.QueryString("QsPageNumber"))
        SkipRec = (PageNumber*RecPerPage)-RecPerPage
    End if

    'response.Write(TotalRec)
    'response.Write(RecPerPage)
    'response.Write(LastPage)

%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="CSS/bootstrap.css">
    <title>Home Expense Tracker - Persons</title>
    <style>
        .wrapper {
            border: 2px solid black;
            background-color: lightgrey;
            width: 20%;
            margin: 0px auto;
            text-align: center;
            height: max-content;
            padding: 10px;
            border-radius: 4px;
        }

        .wrapper label {
            font-size: 18px;
        }

        .combo-account {
            align-content: center;
            font-weight: 700;
        }

        .btn-select {
            background-color: rgb(165, 178, 192);
            border: 1px solid black;
            border-radius: 10px;
            text-align: center;
            padding: 5px;
            font-weight: 700;
            width: 15%;
            cursor: pointer;
            width: 50%;
        }

        .btn-select:hover {
            color: white;
            border: 1px solid whitesmoke;
        }

        .form-control {
            width: 80%;
            margin: auto;
        }

        .table-wrapper {
            background-color: white;
            margin: 10px auto;
            margin-top: 30px;
            width: 40%;

        }

        .table-wrapper th {
            text-align: center;
        }

        .page-nav {
            width: 20%;
            margin: 0px auto;
            text-align: center;
            height: max-content;
            padding: 10px;
        }

        .btn-primary {
            padding: 8px 20px;
        }

        .icon {
            width: 25px;
            height: 25px;
        }
    </style>
</head>

<body>
    <!--- #include file="Header.asp" -->

    <div class="container-fluid">
        <div class="row">
            <div class="col">
                <header class="text-center">
                    <h1>Persons</h1>
                </header>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="wrapper">
                    <form action="Persons.asp" method="post" class="combo-account">
                        <div class="form-group">
                            <input type="hidden" name="TxtPersId" value="<% response.write(PersonId) %>" />
                            <label>Person Name</label>
                            <input type="text" name="TxtPersnName" value="<% response.write(PersonName) %>"
                                class="form-control" />
                            <br>
                            <% if request.QueryString("Action")="1" then %>
                            <input type="submit" value="Update Person" class="btn-select" />
                            <% else %>
                            <input type="submit" value="Save Person" class="btn-select" />
                            <% end if %>
                        </div>
                    </form>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="table-wrapper">
                    <table class="table table-bordered table-hover">
                        <thead class="thead-light">
                            <tr>
                                <th style="width: 8%;">Person Id</th>
                                <th style="width: 12%;">Person Name</th>
                                <th style="width: 5%;"></th>
                                <th style="width: 5%;"></th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
            
                            Dim RSPerson
                            Set RSPerson = Server.CreateObject("ADODB.RecordSet")
                            RSPerson.Open "SELECT * FROM Persons ORDER BY PersonName ASC",conn
                            
                            Dim SkipCounter
                            SkipCounter=1
                            RecNumber=0
            
                            'response.Write("Skip Rec =" & SkipRec)
                            'response.Write("<br>Page Number =" & PageNumber)
            
                            do while not RSPerson.EOF
            
                            if SkipCounter > SkipRec then
                        %>
                            <tr>
                                <td><% Response.Write(RSPerson("PersonId")) %></td>
                                <td><% Response.write(RSPerson("PersonName")) %></td>
                                <td><a href="Persons.asp?Action=1&QsID=<% response.write(RSPerson("PersonId")) %>"><img
                                            src="images/edit.png" alt="Edit" class="icon"></a></td>
                                <td><a href="Persons.asp?Action=2&QsID=<% response.write(RSPerson("PersonId")) %>"><img
                                            src="images/delete.png" alt="Delete" class="icon"></a></td>
                            </tr>
                            <% 
                            RecNumber = RecNumber+1
                            end if
            
                            if RecPerPage = RecNumber then
                                exit do
                            end if
            
                            skipcounter = skipcounter+1
            
                            RSPerson.MoveNext
                            loop 
            
                            RSPerson.Close
                            Set RSPerson = Nothing
            
                            Conn.Close
                            Set Conn = Nothing
                        %>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="page-nav">
                    <% if LastPage = 0 or PageNumber <=1 then %>
                    <a href="Persons.asp?QsPageNumber=1" class="btn btn-primary disabled">First</a>
                    <% else %>
                    <a href="Persons.asp?QsPageNumber=1" class="btn btn-primary">First</a>
                    <% End if %>

                    <% if pagenumber > 1 then %>
                    <a href="Persons.asp?QsPageNumber=<% response.write(PageNumber-1) %>"
                        class="btn btn-primary ">Previous</a>
                    <% else %>
                    <a href="Persons.asp?QsPageNumber=<% response.write(PageNumber-1) %>"
                        class="btn btn-primary disabled">Previous</a>
                    <% End if %>

                    <% if LastPage > 1 then %>
                    <a href="Persons.asp?QsPageNumber=<% response.write(PageNumber+1) %>"
                        class="btn btn-primary">Next</a>
                    <% else %>
                    <a href="Persons.asp?QsPageNumber=<% response.write(PageNumber+1) %>"
                        class="btn btn-primary disabled">Next</a>
                    <% end if %>

                    <% if LastPage >1 then %>
                    <a href="Persons.asp?QsPageNumber=<% response.write(LastPage) %>" class="btn btn-primary">Last</a>
                    <% else %>
                    <a href="Persons.asp?QsPageNumber=<% response.write(LastPage) %>"
                        class="btn btn-primary disabled">Last</a>
                    <% End if %>
                </div>
            </div>
        </div>
    </div>

    <!--- #include file="footer.asp" -->
</body>

</html>