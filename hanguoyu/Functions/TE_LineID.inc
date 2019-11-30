<script type="text/VBscript" runat="server">
  Function TE_ReLineID(ByVal LineID As string) As boolean
    Dim ReLineID As new regex("^[a-zA-Z0-9_\.\-]{2,20}$")   'LineID¹ï¤ñ
    return ReLineID.IsMatch(LineID)
  end function
</script>