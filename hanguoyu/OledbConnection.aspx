<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%
    '''ASP��k�}�Ҹ�Ʈw�����bpage�ŧiaspcompat="true"(�Φb�s���B�ק�B�R��)
    Dim ConnectionDbFile, ConnectionDbPsw,con

    ConnectionDbPsw = "mynameisleyta1992"
    ConnectionDbFile = Request.PhysicalApplicationPath & "mdb\ASP_Web_Single_advertisement.accdb"

    if My.computer.FileSystem.FileExists(ConnectionDbFile) then con = server.CreateObject("ADODB.Connection")
%>
<script runat="server">
    function ConnectionText(DataSource as string, DbPasswoed as string)
        Dim ConnectionSource = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & DataSource & ";Jet OLEDB:Database Password=" & DbPasswoed & ";"
        return ConnectionSource
    end function
</script>