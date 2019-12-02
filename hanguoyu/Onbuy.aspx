<%@ Page Language="VBScript" Title="付款 RFID個資防盜卡 | 韓粉必備" AutoEventWireup="true" aspcompat=true Debug="true"%>
<!--#include file="Functions/TE_ReturnURL.inc"-->
<!--#include file="Functions/TE_Re0_9.inc"-->
<!--#include file="Functions/TE_ReA_Z_0_9.inc"-->
<!--#include file="OledbConnection.aspx"-->
<!DOCTYPE html>
<html>
<head runat="server">
    <title><%=Page.Title%> - 氣噗噗地下商店</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <link rel=stylesheet type="text/css" href="~/CSS/FontAwesome/all.min.css">
    <link rel=stylesheet type="text/css" href="~/CSS/AppStyle.css">
    <link rel=stylesheet type="text/css" href="~/index_css/index-css-1.css">
    <link rel=stylesheet type="text/css" href="~/Onbuy_css/OnbuyCSS01.css">
    <link rel=stylesheet type="text/css" href="~/Onbuy_css/OnbuyCSS02.css">
    <script type="text/javascript" src="JS/TeAjax.js"></script>
</head>
<body>
<%''' 選單  [動態函數未完成]%>
<!--#include file="MenuUrl.aspx"-->
    <div id="chipupu-index-content">
    <%
        Dim OdF_Array(1,1)  '''動態陣列
        Dim ObF_inputNameList  '''商品表單欄位名稱字串
        Dim Ob_CommText, Ob_CommWhere1, Ob_Execute  '''SQL查詢字串, SQL查詢條件, 執行資料庫

        ObF_inputNameList = Trim(Request.Form("inputNameList"))
        if Len(ObF_inputNameList) > 0 then
            Ob_CommWhere1 = ""
            ObF_inputNameList = ObF_inputNameList.Split(",")
            redim preserve OdF_Array(ObF_inputNameList.length - 1, 1)
            for i = 0 to ObF_inputNameList.length - 1
                if not TE_ReA_Z_0_9(ObF_inputNameList(i)) or not TEReNumber09(Trim(Request.Form(ObF_inputNameList(i)))) then
                    Ob_CommWhere1 = ""
                    Exit for
                else
                    OdF_Array(i,0) = {ObF_inputNameList(i),Trim(Request.Form(ObF_inputNameList(i)))}
                    if Len(Ob_CommWhere1) > 0 then
                        Ob_CommWhere1 &= ",'" & ObF_inputNameList(i) & "'"
                    else
                        Ob_CommWhere1 = "'" & ObF_inputNameList(i) & "'"
                    end if
                end if
            Next

            Dim Od_ArrayIndex  '''陣列索引
            Dim Od_ArrayData2  '''陣列資料2維
            try
                Od_ArrayIndex = 0
                if Len(Ob_CommWhere1) > 0 then
                    con.open(ConnectionText(ConnectionDbFile, ConnectionDbPsw))
                    Ob_CommText = "Select Cod01, Cod02, Cod03, Cod04, Cod06 From Commodity Where Cod06 in (" & Ob_CommWhere1 & ")"
                    Ob_Execute = con.execute(Ob_CommText)
                    if not Ob_Execute.Eof then
%>
                        <div id="Onbuy-Content-view">
                            <div id="Onbuy-bank-data">
                                <h2><i class="fas fa-comments-dollar"></i> 匯款資料</h2>
                                <h4>不好意思，目前只提供轉帳匯款，不便請見諒...</h4>
                                <h4>注意，請先匯款再提交資料喔</h4>
                                <p><span>銀行名稱:合作金庫(代號:006)</span>
                                <p><span>銀行帳號:1070765778320</span>
                                <p><span>帳戶戶名:廖駿勝</span>
                                <div id="Onbuy-form-content">
                                    <div class="onbuyf-input">
                                        <h2><i class="fas fa-box"></i> 收件資料</h2>
                                    </div>
                                    <form id="onbuyf" method="post">
                                        <div class="onbuyf-input">
                                            <p><label for="onbuyf01">收件人</label></p>
                                            <input type="text" id="onbuyf01">
                                        </div>
                                        <div class="onbuyf-input">
                                            <p><label for="onbuyf02">收件地址</label></p>
                                            <input type="text" id="onbuyf02">
                                        </div>
                                        <div class="onbuyf-input">
                                            <p><label for="onbuyf03">手機號碼</label></p>
                                            <input type="text" id="onbuyf03">
                                        </div>
                                        <div class="onbuyf-input">
                                            <p><label for="onbuyf04">電子信箱</label></p>
                                            <input type="text" id="onbuyf04">
                                        </div>
                                        <div class="onbuyf-input">
                                            <p><label for="onbuyf05">匯款金額</label></p>
                                            <input type="number" id="onbuyf05">
                                        </div>
                                        <div class="onbuyf-input">
                                            <p><label for="onbuyf06">匯款後5碼</label></p>
                                            <input type="text" id="onbuyf06">
                                        </div>
                                        <div class="onbuyf-input">
                                            <button type="button">提交</button>
                                        </div>
                                    </form>
                                </div>
                                <div id="onbuyf-Ajax"></div>
                            </div>
                            <div id="Onbuy-Buy-data">
                                <h2><i class="fas fa-barcode"></i> 購買品項</h2>
<%
                                Do while not Ob_Execute.Eof
                                   if OdF_Array(Od_ArrayIndex, 0)(0) = Ob_Execute("Cod06").value then
                                       Od_ArrayData2 = OdF_Array(Od_ArrayIndex, 0)(1)
                                   end if                                    
%>
                                    <div class="Onbuy-commodity-shell">
                                        <div class="Onbuy-commodity-name"><%=Ob_Execute("Cod02").value%></div>
                                        <div class="Onbuy-commodity-quantity"><input type="number" value="<%=Od_ArrayData2%>"></div>
                                        <div class="Onbuy-commodity-money">300</div>
                                    </div>
<%
                                    Ob_Execute.movenext
                                    Od_ArrayIndex += 1
                                Loop
%>
                            </div>
                        </div>
<%
                    end if
                    con.Close()
                else
                    Response.Write("<br><h1 class=""alignCenter"">請回首業重新提交資料系，謝謝...<i class=""fas fa-bomb""></i></h1>")
                end if
            catch err1 as Exception
                Response.Write("<br><h1 class=""alignCenter"">資料提交發現錯誤，請回首頁重新提交...<i class=""fas fa-bomb""></i></h1>")
            end try            
        end if
    %>
    </div>
    <div id="chipupu-index-fooder">
        <div id="chipupu-index-fooder-text">
            氣噗噗地下商店© 2019 版權所有
        </div>
    </div>
</body>
</html>






