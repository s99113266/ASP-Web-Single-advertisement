<%@ Page Language="VBScript" Title="商品設定(RFID個資防盜卡 | 韓粉必備)" AutoEventWireup="true" aspcompat=true%>
<!--#include file="~/OledbConnection.aspx"-->

<%
    con.open(ConnectionText(ConnectionDbFile, ConnectionDbPsw))



    con.close()
%>

<!DOCTYPE html>
<html>
<head runat="server">
    <title><%=Page.Title%> - 氣噗噗地下商店</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel=stylesheet type="text/css" href="~/CSS/FontAwesome/all.min.css">
    <link rel=stylesheet type="text/css" href="~/CSS/AppStyle.css">
    <link rel=stylesheet type="text/css" href="~/index_css/index-css-1.css">
    <link rel=stylesheet type="text/css" href="Z0_index_CSS/Z0_INDEXCSS01.css">
    <script type="text/javascript" src="../JS/TeAjax.js"></script>
    <script type="text/javascript" src="../JS/TeFormReset.js"></script>
</head>
<body>
<%''' 選單  [動態函數未完成]%>
<!--#include file="Z0_MenuUrl.aspx"-->

    <div id="chipupu-index-content">
        <div id="chipupu-Z0-index-formBody">
            <form id="z0f" method="post" enctype="multipart/form-data">
                <table>
                    <tr>
                        <td class="Z0-index-formBody-Title"><label for="z0f1">商品名稱</label></td>
                        <td><input type="text" id="z0f1" name="z0f1"></td>
                    <tr>
                        <td class="Z0-index-formBody-Title"><label for="z0f2">商品金額</label></td>
                        <td><input type="number" id="z0f2" name="z0f2" value="0"></td>
                    <tr>
                        <td class="Z0-index-formBody-Title"><label for="z0f3">商品圖片</label></td>
                        <td><input type="file" id="z0f3" name="z0f3" accept="image/*" enctype="multipart/form-data" multiple="multiple"></td>
                    <tr>
                        <td class="Z0-index-formBody-Title Z0-index-formBody-Title-top"><label for="z0f4">商品介紹</label></td>
                        <td><textarea id="z0f4" name="z0f4"></textarea></td>
                    <tr>
                        <td></td>
                        <td>
                            <button type="button" id="Z0-index-form-but1">新增</button>
                        </td>
                </table>
            </form>
        </div>
        <div id="chipupu-Z0-index-form-"></div>
    </div>
    <div id="chipupu-index-fooder">
        <div id="chipupu-index-fooder-text">
            氣噗噗地下商店© 2019 版權所有
        </div>
    </div>
</body>
</html>
<script>
    window.addEventListener("load", async function(){
        var z0ClickBut1;
        z0ClickBut1 = document.getElementById("Z0-index-form-but1");
        if(z0ClickBut1){
            z0ClickBut1.addEventListener("click",async function(){
                await TEAjaxFormData(
                    this,
                    'z0f',
                    'Z0/Z0F01.aspx',
                    'POST',
                    'Z0F01'
                ).then(
                    function(z0fon){
                        let FormErr;
                        FormErr = "";
                        switch (z0fon){
                            case "1":
                                FormErr = "商品新增完成!";
                                FormReset("chipupu-Z0-index-formBody");
                                break;
                            default:
                                FormErr = z0fon;
                                break;
                        }
                        alert(FormErr);
                    },
                    function(z0foff){alert(z0foff);}
                )
            })
        }
    });
</script>




