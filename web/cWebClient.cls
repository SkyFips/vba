VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWebClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' License:     MIT (https://opensource.org/license/mit/)                        |
' Versioning:  https://semver.org                                               |
' Source:      https://github.com/SkyFips/vba/tree/main/web                     |
' Description: cWebClient executes requests and handles response and is         |
'              responsible for functionality shared between requests,           |
'              such as authentication, proxy configuration, and security.       |
'              forked from https://github.com/VBA-tools/VBA-Web                 |
'              based on winhttprequest                                          |
'              learn.microsoft.com/en-us/windows/win32/winhttp/winhttprequest   |
' ------------------------------------------------------------------------------|
' Version | Description                                                         |
'   0.1.0 | initial version                                                     |
' ##############################################################################/
Option Explicit
Dim oLogger As iLogger
Private Const web_HttpRequest_SetCredentials_ForServer = 0
Private Const web_HttpRequest_SetCredentials_ForProxy = 1
Private Const web_HttpRequest_ProxySetting_Default = 0
Private Const web_HttpRequest_ProxySetting_PreConfig = 0
Private Const web_HttpRequest_ProxySetting_Direct = 1
Private Const web_HttpRequest_ProxySetting_Proxy = 2

' WinHttpRequestOption
' https://learn.microsoft.com/en-us/windows/win32/winhttp/winhttprequestoption
Private Enum winHttpRequestOption
  userAgentString = 0
  URL = 1
  URLCodePage = 2
  escapePercentInURL = 3
  sslErrorIgnoreFlags = 4
  selectCertificate = 5
  enableRedirects = 6
  urlEscapeDisable = 7
  urlEscapeDisableQuery = 8
  secureProtocols = 9
  enableTracing = 10
  revertImpersonationOverSsl = 11
  enableHttpsToHttpRedirects = 12
  enablePassportAuthentication = 13
  maxAutomaticRedirects = 14
  maxResponseHeaderSize = 15
  maxResponseDrainSize = 16
  enableHttp1_1 = 17
  enableCertificateRevocationCheck = 18
End Enum

Private web_pProxyServer     As String
Private web_pAutoProxyDomain As String
Public baseUrl               As String
Public authenticator         As iWebAuthenticator
Public proxyBypassList       As String
Public proxyUsername         As String
Public ProxyPassword         As String
Public enableAutoProxy       As Boolean
Public insecure              As Boolean
Public followRedirects       As Boolean
Public timeOutMs             As Long

Private Sub Class_Initialize()
  Me.timeOutMs = 50000
  Me.enableAutoProxy = False
  Me.insecure = False
  Me.followRedirects = True
End Sub
Public Property Set logger(l As iLogger)
  Set oLogger = l
End Property
Public Property Get logger() As iLogger
  If oLogger Is Nothing Then Set oLogger = New cLogger_watchWindow
  logger = oLogger
End Property
Public Property Get proxyServer() As String
  proxyServer = web_pProxyServer
End Property
Public Property Let proxyServer(Value As String)
  Me.enableAutoProxy = False
  web_pProxyServer = Value
End Property

Public Function execute(Request As WebRequest) As WebResponse
  Dim web_Http As Object
  Set execute = New WebResponse
    
  On Error GoTo web_ErrorHandling

  Set web_Http = prepareHttpRequest(Request)

  Call web_Http.Send(Request.Body)
  Do While Not web_Http.WaitForResponse(0.025)
    VBA.DoEvents
  Loop

  Call execute.CreateFromHttp(client:=Me, Request:=Request, Http:=web_Http)
  Call WebHelpers.LogResponse(client:=Me, Request:=Request, Response:=execute)

  If Not Me.authenticator Is Nothing Then Call Me.authenticator.AfterExecute(client:=Me, Request:=Request, Response:=execute)

  Set web_Http = Nothing
  Exit Function

web_ErrorHandling:

  Set web_Http = Nothing
  Dim web_ErrorDescription As String
  Dim web_KeyValue As Object
  
  Set web_KeyValue = CreateObject("Scripting.Dictionary")
  web_ErrorDescription = "An error occurred during execute" & vbNewLine & _
                         Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.description

  WebHelpers.LogError web_ErrorDescription, "cWebClient.Execute", 11011 + vbObjectError

  logger.log DEBUGGER, "<-- Response - " & Format(Now, "Long Time")
  logger.log DEBUGGER, execute.StatusCode & " " & execute.StatusDescription

  For Each web_KeyValue In execute.Headers
    logger.log DEBUGGER, web_KeyValue("Key") & ": " & web_KeyValue("Value")
  Next web_KeyValue

  For Each web_KeyValue In execute.Cookies
    logger.log DEBUGGER, "Cookie: " & web_KeyValue("Key") & "=" & web_KeyValue("Value")
  Next web_KeyValue

  logger.log DEBUGGER, ""
  logger.log DEBUGGER, execute.Content
  logger.log DEBUGGER, ""

  Err.Raise 11011 + vbObjectError, "cWebClient.Execute", web_ErrorDescription
End Function

Public Sub SetProxy(proxyServer As String, _
                    Optional Username As String = "", _
                    Optional password As String = "", _
                    Optional BypassList As String = "")
  Me.proxyServer = proxyServer
  Me.proxyUsername = Username
  Me.ProxyPassword = password
  Me.proxyBypassList = BypassList
End Sub

Public Function GetFullUrl(Request As WebRequest) As String
  GetFullUrl = WebHelpers.JoinUrl(Me.baseUrl, Request.FormattedResource)
End Function

Private Function prepareHttpRequest(Request As WebRequest, Optional Async As Boolean = True) As Object
  Dim web_Http     As Object
  Dim web_KeyValue As Object

  On Error GoTo web_ErrorHandling
    
  Set web_Http = CreateObject("WinHttp.WinHttpRequest.5.1")
  Set web_KeyValue = CreateObject("Scripting.Dictionary")

  ' Prepare request (before open)
  Call web_BeforeExecute(Request)
    
  ' Open http request
  Call web_Http.Open(Method:=WebHelpers.MethodToName(Request.Method), URL:=Me.GetFullUrl(Request), Async:=Async)

  ' Set timeouts
  Call web_Http.SetTimeouts(ResolveTimeout:=Me.timeOutMs, _
                            ConnectTimeout:=Me.timeOutMs, _
                            SendTimeout:=Me.timeOutMs, _
                            ReceiveTimeout:=Me.timeOutMs)

  ' Load auto-proxy (if needed)
  If Me.enableAutoProxy Then Call web_LoadAutoProxy(Request)

  ' Setup proxy
  ' See http://msdn.microsoft.com/en-us/library/windows/desktop/aa384059(v=vs.85).aspx for details
  If Me.proxyServer = vbNullString Then
    WebHelpers.LogDebug "SetProxy: " & Me.proxyServer, "cWebClient.PrepareHttpRequest"
    web_Http.SetProxy web_HttpRequest_ProxySetting_Proxy, Me.proxyServer, Me.proxyBypassList

    If Me.proxyUsername <> vbNullString Then
      WebHelpers.LogDebug "SetProxyCredentials: " & Me.proxyUsername & ", " & WebHelpers.Obfuscate(Me.ProxyPassword), "cWebClient.PrepareHttpRequest"
      web_Http.SetCredentials Me.proxyUsername, Me.ProxyPassword, web_HttpRequest_SetCredentials_ForProxy
    End If
  Else
    ' Attempt to get proxy setup with Proxycfg.exe, otherwise direct
    web_Http.SetProxy web_HttpRequest_ProxySetting_PreConfig
  End If

  ' Setup security
  If Me.insecure Then
    ' - Disable certifcate revocation check
    ' - Ignore all SSL errors
    '   Unknown certification authority (CA) or untrusted root, 0x0100
    '   Wrong usage, 0x0200
    '   Invalid common name (CN), 0x1000
    '   Invalid date or certificate expired, 0x2000
    '   = 0x3300 = 13056
    ' - Enable https-to-http redirects
    web_Http.Option(winHttpRequestOption.enableCertificateRevocationCheck) = False
    web_Http.Option(winHttpRequestOption.sslErrorIgnoreFlags) = 13056
    web_Http.Option(winHttpRequestOption.enableHttpsToHttpRedirects) = True
  Else
    ' By default:
    ' - Enable certificate revocation check (especially useful after HeartBleed)
    ' - Ignore no SLL erros
    ' - Disable https-to-http redirects
    web_Http.Option(winHttpRequestOption.enableCertificateRevocationCheck) = True
    web_Http.Option(winHttpRequestOption.sslErrorIgnoreFlags) = 0
    web_Http.Option(winHttpRequestOption.enableHttpsToHttpRedirects) = False
  End If

  ' Setup redirects
  web_Http.Option(winHttpRequestOption.enableRedirects) = Me.followRedirects

  ' Set headers on http request (after open)
  For Each web_KeyValue In Request.Headers
    Call web_Http.SetRequestHeader(header:=web_KeyValue("Key"), Value:=web_KeyValue("Value"))
  Next web_KeyValue

  For Each web_KeyValue In Request.Cookies
    web_Http.SetRequestHeader "Cookie", web_KeyValue("Key") & "=" & web_KeyValue("Value")
  Next web_KeyValue

  ' Give authenticator opportunity to update Http
  If Not Me.authenticator Is Nothing Then
    Call Me.authenticator.PrepareHttp(client:=Me, Request:=Request, Http:=web_Http)
  End If

  logger.log DEBUGGER, "--> Request - " & Format(Now, "Long Time")
  logger.log DEBUGGER, MethodToName(Request.Method) & " " & client.GetFullUrl(Request)
  
  For Each web_KeyValue In Request.Headers
    logger.log DEBUGGER, web_KeyValue("Key") & ": " & web_KeyValue("Value")
  Next web_KeyValue

  For Each web_KeyValue In Request.Cookies
    logger.log DEBUGGER, "Cookie: " & web_KeyValue("Key") & "=" & web_KeyValue("Value")
  Next web_KeyValue

  If Not IsEmpty(Request.Body) Then
    logger.log DEBUGGER, CStr(Request.Body)
  End If
  Set prepareHttpRequest = web_Http
  Exit Function

web_ErrorHandling:

  Set web_Http = Nothing
  Err.Raise 11012 + vbObjectError, "cWebClient.PrepareHttpRequest", _
            "An error occurred while preparing http request" & vbNewLine & _
            Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.description
End Function

Public Function PrepareCurlRequest(Request As WebRequest) As String
  Dim web_Curl         As String
  Dim web_KeyValue     As Dictionary
  Dim web_CookieString As String

  On Error GoTo web_ErrorHandling
    
  web_Curl = "curl -i"
    
  ' Setup authenticator
  web_BeforeExecute Request

  ' Set timeouts
  ' (max time = resolve + sent + receive)
  web_Curl = web_Curl & " --connect-timeout " & Me.timeOutMs / 1000
  web_Curl = web_Curl & " --max-time " & 3 * Me.timeOutMs / 1000

  ' Setup proxy
  If Me.proxyServer <> "" Then
    web_Curl = web_Curl & " --proxy " & Me.proxyServer

    If Me.proxyBypassList <> "" Then
      web_Curl = web_Curl & " --noproxy " & Me.proxyBypassList
    End If
    If Me.proxyUsername <> "" Then
      web_Curl = web_Curl & " --proxy-user " & Me.proxyUsername & ":" & Me.ProxyPassword
    End If
  End If

  ' Setup security
  If Me.insecure Then
    web_Curl = web_Curl & " --insecure"
  End If

  ' Setup redirects
  If Me.followRedirects Then
    web_Curl = web_Curl & " --location"
  End If
    
  ' Enable compressed if Accept-Encoding != "identity"
  If Request.AcceptEncoding <> "identity" Then
    web_Curl = web_Curl & " --compressed"
  End If

  ' Set headers and cookies
  For Each web_KeyValue In Request.Headers
    web_Curl = web_Curl & " -H '" & web_KeyValue("Key") & ": " & web_KeyValue("Value") & "'"
  Next web_KeyValue

  For Each web_KeyValue In Request.Cookies
    web_CookieString = web_CookieString & web_KeyValue("Key") & "=" & web_KeyValue("Value") & ";"
  Next web_KeyValue
  If web_CookieString <> "" Then
    web_Curl = web_Curl & " --cookie '" & web_CookieString & "'"
  End If

  ' Add method, data, and url
  web_Curl = web_Curl & " -X " & WebHelpers.MethodToName(Request.Method)
  web_Curl = web_Curl & " -d '" & Request.Body & "'"
  web_Curl = web_Curl & " '" & Me.GetFullUrl(Request) & "'"

  ' Give authenticator opportunity to update cURL
  If Not Me.authenticator Is Nothing Then Call Me.authenticator.PrepareCurl(client:=Me, Request:=Request, Curl:=web_Curl)

  ' Log request and return
  Call WebHelpers.LogRequest(client:=Me, Request:=Request)
  PrepareCurlRequest = web_Curl
  Exit Function

web_ErrorHandling:
  Err.Raise 11013 + vbObjectError, "cWebClient.PrepareCURLRequest", _
            "An error occurred while preparing cURL request" & vbNewLine & _
            Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.description
End Function

' Clone client
'
' @internal
' @method Clone
' @return {cWebClient}
Public Function Clone() As cWebClient
  Set Clone = New cWebClient
  Clone.baseUrl = Me.baseUrl
  Clone.proxyServer = Me.proxyServer
  Clone.proxyBypassList = Me.proxyBypassList
  Clone.proxyUsername = Me.proxyUsername
  Clone.ProxyPassword = Me.ProxyPassword
  Clone.enableAutoProxy = Me.enableAutoProxy
  Clone.timeOutMs = Me.timeOutMs
  Clone.insecure = Me.insecure
  Set Clone.authenticator = Me.authenticator
End Function

Private Sub web_BeforeExecute(web_Request As WebRequest)
  If Not Me.authenticator Is Nothing Then Call Me.authenticator.BeforeExecute(client:=Me, Request:=web_Request)

  ' Preparing request includes adding headers
  ' -> Needs to happen after BeforeExecute in case headers were changed
  web_Request.Prepare
End Sub

Private Sub web_LoadAutoProxy(web_Request As WebRequest)
  On Error GoTo web_ErrorHandling

  Dim web_Parts As Dictionary
  Dim web_Domain As String
  Dim web_ProxyServer As String
  Dim web_ProxyBypassList As String

  Set web_Parts = WebHelpers.GetUrlParts(Me.GetFullUrl(web_Request))
  web_Domain = VBA.IIf(web_Parts("Protocol") <> "", web_Parts("Protocol") & "://", "") & _
    web_Parts("Host") & ":" & web_Parts("Port")

  ' Cache auto-proxy by domain
  If web_Domain <> web_pAutoProxyDomain Then
    ' Cache first to store error as no proxy
    web_pAutoProxyDomain = web_Domain

    WebHelpers.GetAutoProxy web_Domain, web_ProxyServer, web_ProxyBypassList

    WebHelpers.LogDebug "Loaded auto-proxy for " & web_Domain & ":" & vbNewLine & _
                        "Server = " & web_ProxyServer & vbNewLine & _
                        "Bypass List = " & web_ProxyBypassList

    ' Store proxy server in underlying to avoid turning off auto-proxy
    web_pProxyServer = web_ProxyServer
    Me.proxyBypassList = web_ProxyBypassList
  End If

  Exit Sub

web_ErrorHandling:
    WebHelpers.LogError "An error occurred while loading auto-proxy" & vbNewLine & _
                        Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": " & Err.description, _
                        "cWebClient.LoadAutoProxy", Err.Number
End Sub
