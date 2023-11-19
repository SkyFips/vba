VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ##############################################################################\
' Author(s):   Philipp Gorkiewicz                                               |
' Description: initially copied from https://github.com/VBA-tools/VBA-Web       |
'              Interface for creating authenticators for rest client            |
'                                                                               |
' ##############################################################################|
'   Version | Description                                                       |
'     1.0.0 | Initial Version                                                   |
' ##############################################################################/

' Hook for taking action before a request is executed
'
' @method BeforeExecute
' @param {cWebClient} Client The client that is about to execute the request
' @param in|out {cWebRequest} Request The request about to be executed
Public Sub BeforeExecute(ByVal client As cWebClient, ByRef Request As cWebRequest)
  ' e.g Add headers, cookies, etc.
End Sub

' Hook for taking action after request has been executed
'
' @method AfterExecute
' @param {cWebClient} Client The client that executed request
' @param {cWebRequest} Request The request that was just executed
' @param in|out {cWebResponse} Response to request
Public Sub AfterExecute(ByVal client As cWebClient, ByVal Request As cWebRequest, ByRef Response As cWebResponse)
  ' e.g. Handle 401 Unauthorized or other issues
End Sub

' Hook for updating http before send
'
' @method PrepareHttp
' @param {cWebClient} Client
' @param {cWebRequest} Request
' @param in|out {WinHttpRequest} Http
Public Sub PrepareHttp(ByVal client As cWebClient, ByVal Request As cWebRequest, ByRef Http As Object)
  ' e.g. Update option, headers, etc.
End Sub

' Hook for updating cURL before send
'
' @method PrepareCurl
' @param {cWebClient} Client
' @param {cWebRequest} Request
' @param in|out {String} Curl
Public Sub PrepareCurl(ByVal client As cWebClient, ByVal Request As cWebRequest, ByRef Curl As String)
  ' e.g. Add flags to cURL
End Sub
