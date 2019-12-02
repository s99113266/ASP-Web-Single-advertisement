<%@ Page Language="VBScript" Title="RFID個資防盜卡 | 韓粉必備" AutoEventWireup="true" aspcompat=true%>
<!--#include file="Functions/TE_ReturnURL.inc"-->
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
    <script type="text/javascript" src="JS/TeAjax.js"></script>
</head>
<body>
<%''' 選單  [動態函數未完成]%>
<!--#include file="MenuUrl.aspx"-->
    <div id="chipupu-index-content">
        <div id="chipupu-index-content-1">
            <div id="chipupu-index-content-1-content1" class="parallel-width">
                <img src="img/hanguoyu.png">
            </div>
        </div>
        <div id="chipupu-index-content-2">
            <div class="chipupu-index-content-parallel parallel-width">
                <div id="chipupu-index-content-2-content1" class="chipupu-index-content-left">
                    <h2>RFID防盜卡</h2>
                    <p>
                        此防盜卡可以<b>保護個資、信用卡等…</b>多種<b>晶片卡</b>或<b>磁卡</b>被人使用RFID、NFC等技術盜取晶片資料。
                    </p>
                </div>
                <div id="chipupu-index-content-2-content2" class="chipupu-index-content-right">
                    <h2>防盜卡可以幫你做到?</h2>
                    <p>
                        增加晶片卡百分之八十的防護，幫你堤防那些科技小偷，更愉快的旅遊，在人群中更自在。
                        <button type="button"><i class="fas fa-running"></i> 馬上了解!</button>
                    </p>
                </div>
            </div>
        </div>
        <div id="chipupu-index-content-3">
            <div id="chipupu-index-content-3-content1" class="parallel-width">
                <h2>RFID防盜活動</h2>
                <p>
                   作為一個忠實韓國瑜粉絲，為了支持韓市長特別印刷了韓粉必備旗艦款，用來收藏也具備多功能，而且市場防盜卡價格動不動就需要高單價，而我們因為是為了支持韓國瑜市長，所以我們只需要100就能擁有旗艦款，現貨數量限量1,000組，賣完為止喔!<br>
                   <span>由於再活動期間現在只要<b>滿8組</b>就可以免運費喔，如果想進一步了解可以往下找到我們的 <i class="fab fa-line"></i> LINE官方帳號。
                </p>
            </div>
        </div>
        <div id="chipupu-index-content-4">
            <div class="chipupu-index-content-parallel parallel-width">
                <div class="chipupu-index-content-parallel-Row">
                    <h2><i class="fas fa-user-shield"></i></h2>
                    <span>資訊安全</span>
                </div>
                <div class="chipupu-index-content-parallel-Row">
                    <h2><i class="fab fa-viadeo"></i></h2>
                    <span>使用廣泛</span>
                </div>
                <div class="chipupu-index-content-parallel-Row">
                    <h2><i class="fas fa-shipping-fast"></i></h2>
                    <span>運送快速</span>
                </div>
            </div>
        </div>
        <div id="chipupu-index-content-5">
            <div class="chipupu-index-content-parallel parallel-width">
                <div class="chipupu-index-content-parallel-Row">
                    <form id="indexf" method="post" action="Onbuy.aspx">
                        <%

                            Dim index_CommText, index_Execute '''資料庫字串, 資料庫內容
                            Dim index_ImageSplit, index_ImageView  '''圖瑱字串切割, 圖片顯示字串
                            Dim index_Url '''網域提取
                            try
                                index_Url = TE_ReturnURL("mdb/CommodityImage/") 
                                con.open(ConnectionText(ConnectionDbFile, ConnectionDbPsw))
                                index_CommText = "Select Cod01, Cod02, Cod03, Cod04, Cod05, Cod06 From Commodity Where Cod07 in (0)"
                                index_Execute = con.execute(index_CommText)                                
                                Do while not index_Execute.Eof
                                    if Len(index_Execute("Cod04").value) > 0 then
                                        index_ImageSplit = index_Execute("Cod04").value.split(",")
                                        index_ImageView = " style=""background-image:url('" & index_Url & index_ImageSplit(0) & "');"""
                                    end if
                        %>
                                    <div class="chipupu-form-input">
                                        <h3><%=index_Execute("Cod02").value%></h3>
                                        <div class="chipupu-form-img" <%=index_ImageView%>>
                                        </div>
                                        <div class="chipupu-form-Title">
                                            <p>單價 NT<%=index_Execute("Cod03").value%></p>
                                            <label>
                                                數量
                                                <input type="number" name="<%=index_Execute("Cod06").value%>" value="0" OnInput="formBuyData('<%=index_Execute("Cod06").value%>',this);">
                                            </label> 
                                        </div>
                                    </div>
                        <%
                                    index_Execute.movenext
                                Loop
                                con.close()
                        %>
                                <div class="chipupu-form-button">
                                    <input type="hidden" name="inputNameList" value="">
                                    <button type="button" onclick="formBuySubmit();">前往匯款 <i class="fas fa-cash-register"></i></button>
                                </div>
                                <script>
                                    function formBuyData(d1,d2){
                                        ///d1=商品編號,d2=購買數量
                                        let formObject, formInputName1;  ///表單物件,表單物件.輸入框
                                        formObject = document.getElementById("indexf");
                                        if(formObject){
                                            formInputName1 = formObject.inputNameList;
                                            formInputName1.value = formInputName1.value.replace(d1+",","");
                                            formInputName1.value = formInputName1.value.replace(","+d1,"");
                                            formInputName1.value = formInputName1.value.replace(d1,"");
                                            if(d2.value != "0" && d2.value != ""){
                                                if(formInputName1.value.length > 0){
                                                    formInputName1.value += "," + d1;
                                                }else{
                                                    formInputName1.value = d1;
                                                }
                                            }
                                        }
                                    }

                                    function formBuySubmit(){
                                        let formObject;
                                        formObject = document.getElementById("indexf");
                                        if(formObject){
                                            if(formObject.inputNameList.value.length > 0){
                                                formObject.submit();
                                            }else{
                                                alert("請選擇商品數量");
                                            }
                                        }
                                    }

                                    window.onload=function(){
                                        let formObject;
                                        formObject = document.getElementById("indexf");
                                        if(formObject){
                                            formObject.reset();
                                        }
                                    }
                                </script>
                        <%
                            catch err1 as exception
                                Response.Write("<h1 class=""alignCenter"">商品資料讀取錯誤...</h1>")
                            end try
                        %>
                    </form>
                </div>
                <div class="alignRight">
                    <h2 class="alignCenter">聯絡資訊</h2>
                    <p class="alignCenter">Line 官方帳號</p>
                    <div class="chipupu-form-img" style="background-image:url('../img/L.png');"></div>
                    <p class="alignCenter">(點圖放大)</p>
                </div>
            </div>
        </div>
    </div>
    <div id="chipupu-index-fooder">
        <div id="chipupu-index-fooder-text">
            氣噗噗地下商店© 2019 版權所有
        </div>
    </div>
</body>
</html>






