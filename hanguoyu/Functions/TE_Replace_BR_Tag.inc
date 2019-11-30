<script type="VBscript" runat="server">
function TEReplace_BRtoVbcrlf(BRtoVbcrlf)
      Dim BRtoVbcrlf_Coder = replace(BRtoVbcrlf, "<br>", vbcrlf)
      return BRtoVbcrlf_Coder
end function

function TEReplace_VbcrlfToBR(VbcrlfToBR)
    Dim VbcrlfToBR_Coder = replace(VbcrlfToBR, vbcrlf, "<br>")
    return VbcrlfToBR_Coder
end function
</script>