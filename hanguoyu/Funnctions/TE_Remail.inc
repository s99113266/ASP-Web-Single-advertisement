<script language="vbscript" runat="server">
Function rgmail(ByVal mail As string) As boolean
  '���q�l�H�c�榡
  Dim emailfun As new regex("^[a-zA-Z0-9_\.\-]*@[0-9a-zA-Z_\.\-]+([\.])([0-9a-zA-Z]+)$")
  return emailfun.IsMatch(mail)
end function
</script>