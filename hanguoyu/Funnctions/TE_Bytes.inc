<Script type="VBscript" Runat="Server">
Function StrByteLen(dd As String)
  Dim ByteLen = 0
  if Len(dd) > 0 then ByteLen = System.Text.Encoding.Default.GetBytes(dd).Length
  Return ByteLen
End Function
</Script>