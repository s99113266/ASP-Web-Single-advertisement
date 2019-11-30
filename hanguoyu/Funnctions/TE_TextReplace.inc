<script type="VBscript" runat="server">
function TETextReplace(TRText)
  Dim TRCoderASCII = replace(TRText,"<","¡Õ")
  TRCoderASCII = replace(TRCoderASCII,">","¡Ö")
  TRCoderASCII = replace(TRCoderASCII,"|","¡U")
  TRCoderASCII = replace(TRCoderASCII,"^","¡s")
  TRCoderASCII = replace(TRCoderASCII,"%","¢H")
  TRCoderASCII = replace(TRCoderASCII,"'","¡¦")
  return TRCoderASCII
end function
</script>