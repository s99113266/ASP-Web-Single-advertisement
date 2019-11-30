<script language="vbscript" runat="server">
'¹ï¤ñ0-9
Function TEReNumber09(StrNumber09) As Boolean
  Dim re09 As New Regex("^[0-9]+$")
  Return re09.IsMatch(StrNumber09)   
end function
</script>