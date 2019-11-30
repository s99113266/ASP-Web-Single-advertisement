<script type="text/VBscript" runat="server">
function TE_TwPhone(twtxt) As boolean
  Dim twphone As new Regex("^([09]{2}[0-9]{8})$")
  return twphone.ismatch(twtxt)
end function
</script>