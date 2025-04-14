# web
Usage:
Dim Client As New WebClient
Client.BaseUrl = "https://www.example.com/api/"
Dim Auth As New HttpBasicAuthenticator
Auth.Setup Username, Password
Set Client.Authenticator = Auth
Dim Request As New WebRequest
Dim Response As WebResponse
Setup WebRequest...
Set Response = Client.Execute(Request)
-> Uses Http Basic authentication and appends Request.Resource to BaseUrl


Errors:
11010 / 80042b02 / -2147210494 - cURL error in Execute
11011 / 80042b03 / -2147210493 - Error in Execute
11012 / 80042b04 / -2147210492 - Error preparing http request
11013 / 80042b05 / -2147210491 - Error preparing cURL request
