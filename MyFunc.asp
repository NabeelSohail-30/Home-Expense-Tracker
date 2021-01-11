<%
    Function Encrypt(str)
    Dim StrData
    Dim Counter
    Counter = 1

    For Counter = 1 To Len(str)
        StrData = StrData & Chr((Asc(Mid(str, Counter, 1)) + 20))
    Next

    Encrypt = StrData
    End Function

    Function Decrypt(str)
    Dim StrData
    Dim Counter
    Counter = 1

    For Counter = 1 To Len(str)
        StrData = StrData & Chr((Asc(Mid(str, Counter, 1)) - 20))
    Next

    Decrypt = StrData
    End Function
%>