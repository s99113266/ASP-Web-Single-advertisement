<script type="text/VBscript" runat="server">
  Function TE_ReA_Z_0_9(ByVal TE_ReA_Z_0_9Text As string) As boolean
    Dim TE_Text As new regex("^[0-9a-zA-Z]+$")
    return TE_Text.IsMatch(TE_ReA_Z_0_9Text)
  end function
</script>