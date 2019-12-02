<script type="text/VBscript" runat="server">
function TE_ReturnURL(ru)
    return Request.Url.Scheme & "://" & Request.Url.Authority & "/" & ru
end function
</script>