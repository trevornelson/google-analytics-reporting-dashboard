Attribute VB_Name = "modMultichannelAPIQueries"
Option Compare Database

Public Sub MCFtest()

Call MCFQueryHandler("3199261", "ya29.1.AADtN_Uifv3kdlfkYIogiZRka3pCBVRhFEOp2iDI_foA_Da0Kt2_wkpeuu0lE1fcJjeeuQ")


End Sub

Public Sub MCFQueryHandler(ByVal c As String, ByVal aT As String)

Dim m As String
Dim dB As Database
Dim rsMCFAPI As Recordset

Set dB = CurrentDb
Set rsMCFAPI = dB.OpenRecordset("tblMCFApiQueries")

If Not (rsMCFAPI.EOF And rsMCFAPI.BOF) Then
    Do Until rsMCFAPI.EOF = True
    
        m = rsMCFAPI("QueryName").Value
        Call MCFQueryBuild(c, aT, m)
        
        rsMCFAPI.MoveNext
    Loop
Else
    MsgBox "There are no records in the recodset"
End If

rsMCFAPI.Close
dB.Close
Set rsMCFAPI = Nothing
Set dB = Nothing

End Sub


Public Sub MCFQueryBuild(ByVal c As String, ByVal aT As String, ByVal m As String)
'Constructs the Google Analytics MultiChannel Feed API Query

Dim dB As Database
Dim rs As Recordset

Set dB = CurrentDb

Dim mcf_URL As String
Dim APIQueryString As String

APIQueryString = DLookup("[QueryString]", "[tblMCFApiQueries]", "[tblMCFApiQueries]![QueryName]= " & "'" & m & "'")

Dim TabN As String
Dim FieldN As String

TabN = DLookup("[TableName]", "[tblMCFApiQueries]", "[QueryName] = " & "'" & m & "'")
FieldN = DLookup("[FieldName]", "[tblMCFApiQueries]", "[QueryName] = " & "'" & m & "'")

Set rs = dB.OpenRecordset(TabN)

If Not (rs.EOF And rs.BOF) Then
    Do Until rs.EOF = True
        
        Dim Sdate As Date
        Dim Edate As Date
        Dim GAStartDate As String
        Dim GAEndDate As String
        
        Debug.Print TabN
        
        Sdate = rs("MonthStartDate").Value
        Edate = rs("MonthEndDate").Value
        
        GAStartDate = Format(Sdate, "yyyy-mm-dd")
        GAEndDate = Format(Edate, "yyyy-mm-dd")
        
        mcf_URL = "https://www.googleapis.com/analytics/v3/data/mcf?ids=ga:" & c & APIQueryString & "&start-date=" & GAStartDate & "&end-date=" & GAEndDate

        Call MCFQuerySend(mcf_URL, aT, m, TabN, FieldN, Edate)
        
        rs.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset. Bummer."
End If
        
rs.Close
dB.Close
Set rs = Nothing
Set dB = Nothing

End Sub
Public Sub MCFQuerySend(ByVal mcf_URL As String, ByVal aT As String, ByVal m As String, ByVal t As String, ByVal f As String, ByVal endD As String)


Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
    objhttp.Open "GET", mcf_URL, False
    objhttp.setRequestHeader "GET", "/analytics/v3/data/mcf/ HTTP/1.1"
    objhttp.setRequestHeader "Host", "googleapis.com"
    objhttp.setRequestHeader "Authorization", "Bearer " & aT
    objhttp.send
        
    Dim QueryResult As String
    QueryResult = objhttp.responseText
    
    Debug.Print QueryResult


Call MCFQueryParse(QueryResult, m, t, f, endD)

End Sub

Public Sub MCFQueryParse(ByVal qR As String, ByVal m As String, ByVal t As String, ByVal f As String, ByVal endD As String)

Dim sD As Long
Dim eD As Long
Dim ElemTag As String
Dim CharDiff As Long
Dim dP1 As Variant

ElemTag = "primitiveValue"

sD = InStrRev(qR, ElemTag)
sD = sD + 17
eD = InStr(sD, qR, "]]")
eD = eD - 2
CharDiff = eD - sD

If CharDiff > 0 Then
    dP1 = Mid(qR, sD, CharDiff)
Else
    Debug.Print "Looks like this datapoint is zero."
    dP1 = 0
End If

Debug.Print dP1


Call MCFQueryImport(dP1, m, t, f, endD)

End Sub

Public Sub MCFQueryImport(ByVal data, ByVal m As String, ByVal t As String, ByVal f As String, ByVal endD As String)

Dim ImportSQL As String
Debug.Print m
Debug.Print f
Debug.Print t


ImportSQL = "UPDATE " & t & " SET " & f & " = '" & data & "' WHERE MonthEndDate = #" & endD & "#"

Debug.Print ImportSQL

Dim dB As Database
Set dB = CurrentDb

dB.Execute ImportSQL

End Sub
