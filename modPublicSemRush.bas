Attribute VB_Name = "modPublicSemRush"
Option Compare Database


Public Sub semrOrganicKeywordsReport(domain As String)

    Dim dB As Database
    Dim rs As Recordset
    Dim rs2 As Recordset
    'Declares relevant tables and database used for sub. Variable rs will be used as the table containing inputs for creating the URL parameters. The sub will output to the rs2 table.
        
        
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblClients")
    Set rs2 = dB.OpenRecordset("tblSemRushKeywords")
    
    Dim xml As MSXML2.XMLHTTP60
    Dim Result As String
    'Declares xml as a html object using XMLHTTP.
    'Also declares result as a string that will contain the result of the HTTP Get.
        
    Dim URL As String
    URL = "http://us.api.semrush.com/?action=report&type=domain_organic&key=9c2dc60a6ef4ea297e55f10322c6f987&export=api&export_columns=Ph,Po,Pp,Pd&domain=" & domain & "&display_sort=tr_desc"
    'Sets URL equal to current date and US version of semRush API. Also sets parameter string to ApiQuery.
     
    Debug.Print URL
     
    Set xml = CreateObject("MSXML2.XMLHTTP.6.0")
    
    With xml
        .Open "Get", URL
        .send
    End With
    'Sends the URL with parameters to the SEMrush API server.
    
    Result = xml.responseText
    'Returns the server's response in the "result" string variable.
    
    Debug.Print Result
        
      
    Dim FSO As FileSystemObject
    Dim TS As TextStream
    Dim Temps As String
    Set FSO = New FileSystemObject
    Set TS = FSO.OpenTextFile("c:\DigitalStrategyDashboardTempFiles\SemRush1.txt", ForWriting, True)
        TS.Write Result
    TS.Close
    Set TS = Nothing
    Set FSO = Nothing
    'Creates a file in c:\DigitalStrategyDashboardTempFiles called SemRush1.txt and writes the server's response (or "result") to the file.
    'Each time this subroutine is run, it will replace the text in SemRush1.txt
    
    Call dB.Execute("DELETE FROM tblSemRushKeywords")
    
    DoCmd.TransferText acImportDelim, "SemRush1 Import Specification", "tblSemRushKeywords", "c:\DigitalStrategyDashboardTempFiles\SemRush1.txt", True
    'Imports SemRush1.txt to tblSemRushData.
    '******ERROR****** this should replace existing table data, but instead it looks like it is appending it.*****ERROR******
    
    
    rs.Close
    dB.Close
    
                  
End Sub

Public Sub semrLostKeywords(domain As String)

    Dim dB As Database
    Dim rs As Recordset
    Dim rs2 As Recordset
    'Declares relevant tables and database used for sub. Variable rs will be used as the table containing inputs for creating the URL parameters. The sub will output to the rs2 table.
        
        
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblClients")
    Set rs2 = dB.OpenRecordset("tblSemRushlost")
    
    Dim xml As MSXML2.XMLHTTP60
    Dim Result As String
    'Declares xml as a html object using XMLHTTP.
    'Also declares result as a string that will contain the result of the HTTP Get.
        
    Dim URL As String
    URL = "http://us.api.semrush.com/?action=report&type=domain_organic&key=9c2dc60a6ef4ea297e55f10322c6f987&export=api&export_columns=Ph,Po,Pp,Pd&domain=" & domain & "&display_sort=tr_desc&display_positions=lost"
    'Sets URL equal to current date and US version of semRush API. Also sets parameter string to ApiQuery.
     
    Debug.Print URL
     
    Set xml = CreateObject("MSXML2.XMLHTTP.6.0")
    
    With xml
        .Open "Get", URL
        .send
    End With
    'Sends the URL with parameters to the SEMrush API server.
    
    Result = xml.responseText
    'Returns the server's response in the "result" string variable.
    
    Debug.Print Result
        
      
    Dim FSO As FileSystemObject
    Dim TS As TextStream
    Dim Temps As String
    Set FSO = New FileSystemObject
    Set TS = FSO.OpenTextFile("c:\DigitalStrategyDashboardTempFiles\SemRush2.txt", ForWriting, True)
        TS.Write Result
    TS.Close
    Set TS = Nothing
    Set FSO = Nothing
    'Creates a file in c:\DigitalStrategyDashboardTempFiles called SemRush1.txt and writes the server's response (or "result") to the file.
    'Each time this subroutine is run, it will replace the text in SemRush1.txt
    
    Call dB.Execute("DELETE FROM tblSemRushLost")
    
    DoCmd.TransferText acImportDelim, "SemRush1 Import Specification", "tblSemRushLost", "c:\DigitalStrategyDashboardTempFiles\SemRush2.txt", True
    'Imports SemRush1.txt to tblSemRushData.
    '******ERROR****** this should replace existing table data, but instead it looks like it is appending it.*****ERROR******
    
    
    rs.Close
    dB.Close

End Sub

Public Sub semrNewKeywords(domain As String)


    Dim dB As Database
    Dim rs As Recordset
    Dim rs2 As Recordset
    'Declares relevant tables and database used for sub. Variable rs will be used as the table containing inputs for creating the URL parameters. The sub will output to the rs2 table.
        
        
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblClients")
    Set rs2 = dB.OpenRecordset("tblSemRushNew")
    
    Dim xml As MSXML2.XMLHTTP60
    Dim Result As String
    'Declares xml as a html object using XMLHTTP.
    'Also declares result as a string that will contain the result of the HTTP Get.
        
    Dim URL As String
    URL = "http://us.api.semrush.com/?action=report&type=domain_organic&key=9c2dc60a6ef4ea297e55f10322c6f987&export=api&export_columns=Ph,Po,Pp,Pd&domain=" & domain & "&display_sort=tr_desc&display_positions=new"
    'Sets URL equal to current date and US version of semRush API. Also sets parameter string to ApiQuery.
     
    Debug.Print URL
     
    Set xml = CreateObject("MSXML2.XMLHTTP.6.0")
    
    With xml
        .Open "Get", URL
        .send
    End With
    'Sends the URL with parameters to the SEMrush API server.
    
    Result = xml.responseText
    'Returns the server's response in the "result" string variable.
    
    Debug.Print Result
        
      
    Dim FSO As FileSystemObject
    Dim TS As TextStream
    Dim Temps As String
    Set FSO = New FileSystemObject
    Set TS = FSO.OpenTextFile("c:\DigitalStrategyDashboardTempFiles\SemRush3.txt", ForWriting, True)
        TS.Write Result
    TS.Close
    Set TS = Nothing
    Set FSO = Nothing
    'Creates a file in c:\DigitalStrategyDashboardTempFiles called SemRush3.txt and writes the server's response (or "result") to the file.
    'Each time this subroutine is run, it will replace the text in SemRush3.txt
    
    Call dB.Execute("DELETE FROM tblSemRushNew")
    
    DoCmd.TransferText acImportDelim, "SemRush1 Import Specification", "tblSemRushNew", "c:\DigitalStrategyDashboardTempFiles\SemRush3.txt", True
    'Imports SemRush3.txt to tblSemRushNew.
      
    
    rs.Close
    dB.Close
    
End Sub

