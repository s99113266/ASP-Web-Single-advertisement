<script runat="server">
function TE_ClientIP()
    Dim ClientIP = ""
    If Request.Servervariables("HTTP_X_FORWARDED_FOR") <> "" Then
        ClientIP = Request.Servervariables("HTTP_X_Forwarded_For")	' ���ϥ� Proxy
    Else
        ClientIP = Request.Servervariables("Remote_Addr")	' �S���ϥ� Proxy
    End If
    return ClientIP
end function
</script>