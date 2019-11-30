<script runat="server">
function TE_ClientIP()
    Dim ClientIP = ""
    If Request.Servervariables("HTTP_X_FORWARDED_FOR") <> "" Then
        ClientIP = Request.Servervariables("HTTP_X_Forwarded_For")	' 有使用 Proxy
    Else
        ClientIP = Request.Servervariables("Remote_Addr")	' 沒有使用 Proxy
    End If
    return ClientIP
end function
</script>