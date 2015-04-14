Attribute VB_Name = "modOauth"
Option Compare Database

Public Sub Oauth()

    Dim IEApp As InternetExplorer
    Dim URL As String
        
    URL = "https://accounts.google.com/o/oauth2/auth?scope=https://www.googleapis.com/auth/analytics.readonly&response_type=code&client_id=945972137560-p4794vkv6ekjvq8g26rmbldtpc8cu754.apps.googleusercontent.com&redirect_uri=urn:ietf:wg:oauth:2.0:oob"
    
    Set IEApp = New InternetExplorer
    IEApp.Visible = True
    
    IEApp.Navigate URL
    
    
    

End Sub

Public Sub OauthAccessToken()

    Dim URL As String
    Dim bodySend As String
    Dim AuthorizationCode As String
    
    AuthorizationCode = InputBox("Paste Google Analytics Access Token Here:") 'User inputs GA Access Token
    
    Dim dB As Database
    Dim rs As Recordset
                    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblOAuthCredentials") 'Opens connection to the GA credentials table
    
       
    dB.Execute "DELETE * FROM tblOAuthCredentials" 'clears tblOAuthCredientials data from any Oauth flow. This table should only every have one record in it.
    
    rs.AddNew
    rs("AuthorizationCode").Value = AuthorizationCode
    rs.Update
    
    Dim AC As String
    AC = DFirst("AuthorizationCode", "tblOAuthCredentials")
    
    Debug.Print AC
    
    bodySend = "code=" & AC & _
               "&client_id=945972137560-p4794vkv6ekjvq8g26rmbldtpc8cu754.apps.googleusercontent.com" & _
               "&client_secret=PgdtOyHez2f0Heny4iEy89RR" & _
               "&redirect_uri=urn:ietf:wg:oauth:2.0:oob" & _
               "&grant_type=authorization_code"
                   
    
    URL = "https://accounts.google.com/o/oauth2/token"
    
    Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    objhttp.Open "POST", URL, False
    objhttp.setRequestHeader "POST", "/o/oauth2/token HTTP/1.1"
    objhttp.setRequestHeader "Host", "accounts.google.com"
    objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objhttp.send (bodySend)
       
    Dim authResponse As String
    Dim authTokenStart As Integer
        
    authResponse = objhttp.responseText
    Debug.Print authResponse
    
    Dim objJson As MSScriptControl.ScriptControl
    Set objJson = New MSScriptControl.ScriptControl
    objJson.Language = "JScript"
 
    Dim objResp As Object
    Set objResp = objJson.Eval("(" & objhttp.responseText & ")")
    
    Debug.Print objResp
    Debug.Print "right here!"
    Debug.Print objResp.access_token
    
    Dim AccessToken As String
    AccessToken = objResp.access_token
    
    Debug.Print AccessToken
    
    Dim RefreshToken As String
    RefreshToken = objResp.refresh_token

    Debug.Print RefreshToken
    
    Forms!frmDashboard.OauthToken.Value = AccessToken

End Sub

Public Sub OauthRefreshToken()

    'Is just a copy of the oauthaccesstoken sub for the time being.
    
    Dim URL As String
    Dim bodySend As String
    
    bodySend = "code=4/WV9iscGu3odCPWA0otn-HM5QqogV.QoMmc4xM8MseaDn_6y0ZQNjSxMhqfwI" & _
               "&client_id=945972137560.apps.googleusercontent.com" & _
               "&client_secret=9iRnpKMVjzs6dvbFJTkCpd3r" & _
               "&redirect_uri=urn:ietf:wg:oauth:2.0:oob" & _
               "&grant_type=authorization_code"
                   
    
    URL = "https://accounts.google.com/o/oauth2/token"
    
    Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
    
    objhttp.Open "POST", URL, False
    objhttp.setRequestHeader "POST", "/o/oauth2/token HTTP/1.1"
    objhttp.setRequestHeader "Host", "accounts.google.com"
    objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objhttp.send (bodySend)
       
    Dim authResponse As String
    Dim authTokenStart As Integer
    
    authResponse = objhttp.responseText
    Debug.Print authResponse
    Google_RefreshToken = authResponse

End Sub



