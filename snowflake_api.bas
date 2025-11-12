''''''''''''''''''''''''''''''
''Chat GPT Code For REST API''
''''''''''''''''''''''''''''''

Sub RunSnowflakeQuery()

    API_KEY = "vistra_admin_password"

    Set Http = CreateObject("MSXML2.XMLHTTP")

    url = "https://app.us-east-1.privatelink.snowflakecomputing.com/api/v2/statements"

    token = "Bearer " & API_KEY
    
    payload = "{""statement"":" & snowflake_query & "," & _
              """resultSetMetaData"":{""format"":""json""," & _
              """formatOptions"":{""rowLimit"":1000}}}"

    With Http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", token
        .Send payload
        Debug.Print .responseText
    End With

End Sub

Function snowflake_query()
    snowflake_query = "SELECT B.COMMUNITYNAME FROM RETAIL_PRD.LANDPOWER_ODS.PRODUCT_UTILITY_COMMUNITYCONTRACT A" & vbNewLine & _
                        "INNER JOIN RETAIL_PRD.LANDPOWER_ODS.PRODUCT_UTILITY_COMMUNITY B" & vbNewLine & _
                        "ON A.COMMUNITYROWID=B.COMMUNITYROWID" & vbNewLine & _
                        "WHERE A.EXTERNALCONTRACTID='C-00137237'"
    snowflake_query = """" & snowflake_query & """"
End Function

''''''''''''''''''''''''''''''''''''''''''''''
''Ethan's Code For ODBC Driver ADODB Queries''
''''''''''''''''''''''''''''''''''''''''''''''

Sub SnowflakePLSync()

    Dim strCon, strToday, str1Yr As String
    Dim Con As Object, rec As Object
    Dim dtToday, dt1Yr As Date
    
    
    dtToday = Date
    dt1Yr = Date - 368
    
    str1Yr = format(dt1Yr, "YYYY-MM-DD")
    strLog = "Snowflake PL Started " & Now()
    A = FreeFile
    Open logPath For Append As #A
    Print #A, strLog
    Close #A
    
    strCon = "DSN=Snowflake;UID=[USERNAME_HERE];PWD=[PASSWORD_HERE];WAREHOUSE=ADHOC_PRD"
    Set Con = New ADODB.Connection
    Con.ConnectionString = strCon
    Con.CommandTimeout = 660
    Con.Open
    
    strLog = "Snowflake PL Database Connected " & Now()
    A = FreeFile
    Open logPath For Append As #A
    Print #A, strLog
    Close #A
    
    strquery = "select OPRDATE, OPRHOUR, QUANTITY from USERDB_LBM.VW_PJM_DERSR4_A2_SETTLEMENTS where NODENAME = 'DERSR4.DEOK_RESID_AGG-SYSLOAD' and OPRDATE >='"
    strquery = strquery & str1Yr
    strquery = strquery & "' order by OPRDATE, OPRHOUR"
    
    Set rec = CreateObject("ADODB.Recordset")
    
    
    With rec
        .ActiveConnection = Con
        Con.CommandTimeout = 0
        .Open strquery
    End With
    
    For n = 0 To rec.Fields.count - 1
         wsPL.Cells(1, n + 1).value = rec.Fields(n).name
    Next n
    
    strLog = "Snowflake PL Query Complete " & Now()
    A = FreeFile
    Open logPath For Append As #A
    Print #A, strLog
    Close #A
    wsPL.Cells(2, 1).CopyFromRecordset rec
    
    rec.Close
    Set rec = Nothing
    Con.Close
    Set Con = Nothing

End Sub
