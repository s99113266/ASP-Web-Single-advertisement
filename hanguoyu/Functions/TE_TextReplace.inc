<script type="VBscript" runat="server">
function TETextReplace(TRText)
  Dim TRCoderASCII = replace(TRText,"<","��")
  TRCoderASCII = replace(TRCoderASCII,">","��")
  TRCoderASCII = replace(TRCoderASCII,"|","�U")
  TRCoderASCII = replace(TRCoderASCII,"^","�s")
  TRCoderASCII = replace(TRCoderASCII,"%","�H")
  TRCoderASCII = replace(TRCoderASCII,"'","��")
  return TRCoderASCII
end function
</script>