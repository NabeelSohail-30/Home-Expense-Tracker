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

    'Del Catg Start
        If request.QueryString("Action")= "2" AND Request.QueryString("QsId") <> "" then
            QryStr = "DELETE FROM Categories WHERE (CategoryId = " & cint(Request.querystring("QsId")) & ")" 
            'response.Write qrystr
            Conn.Execute QryStr
        end if
    'Del Catg End

    'Save new Catg Start
        If request.Form("CatgName") <> "" And Request.Form("TxtCatgId") = "" then
            'response.Write("Form Posted with value")
            
            QryStr="INSERT INTO Categories (Category) Values('" & Request.Form("CatgName") & "')"
            'response.Write qrystr
            
            Conn.Execute QryStr
        end if
    'Save New Catg End

    'Edit Catg Start
        If request.Form("CatgName") <> "" AND Request.form("TxtCatgId") <> "" then
            'response.Write("Form Posted with value")
            
            QryStr="UPDATE Categories SET Category = '" & Request.Form("CatgName") & "' WHERE (CategoryId = " & cint(Request.form("TxtCatgId")) & ")"
            'response.Write qrystr
            
            Conn.Execute QryStr
        end if
    'Edit Catg End

    'Get Value for Edit
        If request.QueryString("Action")= "1" AND Request.QueryString("QsId") <> "" then
            Dim RSEditCatg
            Set RSEditCatg = Server.CreateObject("ADODB.RecordSet")
            
            QryStr = "SELECT * FROM Categories WHERE (CategoryId = " & cint(Request.QueryString("QsId")) & ")"
            'response.Write qrystr

            RSEditCatg.Open QryStr,Conn
      
            Dim CategoryName
            Dim CatgId
          
            CategoryName = RSEditCatg("Category")
            CatgId = RSEditCatg("CategoryId")  
    
            'response.write CategoryName  

            RSEditCatg.Close
            Set RSEditCatg = Nothing
        end if

    

    Dim RSCount
    Set RSCount = Server.CreateObject("ADODB.RecordSet")
    
    RSCount.Open "SELECT COUNT(CategoryID) AS TotalRecords FROM Categories",Conn

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
    <title>Home Expense Tracker - Categories</title>
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
                    <h1>Categories</h1>
                </header>
            </div>
        </div>

        <div class="row">
            <div class="col">
                <div class="wrapper">
                    <form action="Categories.asp" method="POST" class="combo-account">
                        <div class="form-group">
                            <input type="hidden" name="TxtCatgId" value="<% response.write(CatgId) %>" />
                            <label>Category Name</label>
                            <input type="text" name="CatgName" value="<% response.write(CategoryName) %>"
                                class="form-control" />
                            <br>
                            <% if request.QueryString("Action")="1" then %>
                            <input type="submit" value="Update Category" class="btn-select" />
                            <% else %>
                            <input type="submit" value="Save Category" class="btn-select" />
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
                                <th style="width: 8%;">Category Id</th>
                                <th style="width: 12%;">Category Name</th>
                                <th style="width: 5%;"></th>
                                <th style="width: 5%;"></th>
                            </tr>
                        </thead>

                        <tbody>
                            <%
            
                            Dim RSCatg
                            Set RSCatg = Server.CreateObject("ADODB.RecordSet")
                            RSCatg.Open "SELECT * FROM Categories ORDER BY Category ASC",conn
                            
                            Dim SkipCounter
                            SkipCounter=1
                            RecNumber=0
            
                            'response.Write("Skip Rec =" & SkipRec)
                            'response.Write("<br>Page Number =" & PageNumber)
            
                            do while not RSCatg.EOF
            
                            if SkipCounter > SkipRec then
                        %>
                            <tr>
                                <td><% Response.Write(RSCatg("CategoryID")) %></td>
                                <td><% Response.write(RSCatg("Category")) %></td>
                                <td><a href="Categories.asp?Action=1&QsID=<% response.write(RSCatg("CategoryID")) %>"><img
                                            src="images/edit.png" alt="Edit" class="icon"></a></td>
                                <td><a href="Categories.asp?Action=2&QsID=<% response.write(RSCatg("CategoryID")) %>"><img
                                            src="images/delete.png" alt="Delete" class="icon"></a></td>
                            </tr>
                            <% 
                            RecNumber = RecNumber+1
                            end if
            
                            if RecPerPage = RecNumber then
                                exit do
                            end if
            
                            skipcounter = skipcounter+1
            
                            RSCatg.MoveNext
                            loop 
            
                            RSCatg.Close
                            Set RSCatg = Nothing
            
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
            <a href="Categories.asp?QsPageNumber=1" class="btn btn-primary">First</a>
            <% else %>
            <a href="Categories.asp?QsPageNumber=1" class="btn btn-primary">First</a>
            <% End if %>

            <% if pagenumber > 1 then %>
            <a href="Categories.asp?QsPageNumber=<% response.write(PageNumber-1) %>" class="btn btn-primary ">Previous</a>
            <% else %>
            <a href="Categories.asp?QsPageNumber=<% response.write(PageNumber-1) %>" class="btn btn-primary disabled">Previous</a>
            <% End if %>

            <% if LastPage > 1 then %>
                <a href="Categories.asp?QsPageNumber=<% response.write(PageNumber+1) %>" class="btn btn-primary">Next</a>
            <% else %>
                <a href="Categories.asp?QsPageNumber=<% response.write(PageNumber+1) %>" class="btn btn-primary disabled">Next</a>
            <% end if %>

            <% if LastPage >1 then %>
                <a href="Categories.asp?QsPageNumber=<% response.write(LastPage) %>" class="btn btn-primary">Last</a>
            <% else %>
                <a href="Categories.asp?QsPageNumber=<% response.write(LastPage) %>" class="btn btn-primary disabled">Last</a>
            <% End if %>
                </div>
            </div>
        </div>
    </div>

    <!--- #include file="footer.asp" -->
</body>

</html>