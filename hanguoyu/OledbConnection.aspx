<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%
    '''ASP方法開啟資料庫必須在page宣告aspcompat="true"(用在新曾、修改、刪除)
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