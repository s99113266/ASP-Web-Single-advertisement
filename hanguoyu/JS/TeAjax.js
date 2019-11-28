function TEAjax(TEPageID, TEDataPage) {
    var TExmlhttp, TEURL, TEPageDoc;
    TEURL = location.protocol + "//" + location.host + "/";
    TEPageDoc = document.getElementById(TEPageID);
    return new Promise(function (resolve, reject) {
        if (TEPageDoc) {
            TEPageDoc.innerHTML = "<span id=\"AjaxLoading\"><i class=\"fas fa-sync-alt fa-spin\"></i></span>";
            if (window.XMLHttpRequest) {
                //  IE7+, Firefox, Chrome, Opera, Safari 浏览器执行代码
                TExmlhttp = new XMLHttpRequest();
            } else {
                // IE6, IE5 浏览器执行代码
                TExmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            TExmlhttp.onreadystatechange = function () {
                if (TExmlhttp.readyState == 3 && TExmlhttp.status == 200) {
                    TEPageDoc.innerHTML = "<span id=\"AjaxLoading\"><i class=\"fas fa-sync-alt fa-spin\"></i></span>";
                }
                if (TExmlhttp.readyState == 4 && TExmlhttp.status == 200) {
                    TEPageDoc.innerHTML = TExmlhttp.responseText;
                    resolve("200");
                } else if (TExmlhttp.status == 404) {
                    reject("404");
                } else if (TExmlhttp.status == 401) {
                    reject("401");
                } else if (TExmlhttp.status == 500) {
                    reject("500");
                }
            }
            TExmlhttp.open("POST", TEURL + TEDataPage, true);
            TExmlhttp.send();
        } else {
            reject("Error:404-n");
        }
    });
}

function TEAjaxFormData(TeFormButtonThis,TeFormDataID, TeFormDataURL, TeFormDataMethod, TeFormFilePAGE) {
    var newFormData = document.getElementById(TeFormDataID);
    return new Promise(function (resolve, reject) {
        let ButtonInnerHTML = TeFormButtonThis.innerHTML;
        TeFormButtonThis.disabled = true;
        TeFormButtonThis.innerHTML = "<i class=\"fas fa-spinner fa-pulse\"><\/i>";
        if (newFormData) {
            let TeFormHttp = new XMLHttpRequest();
            let TeFormUrl = location.protocol + "//" + location.host + "/";
            var TeFormError = TeFormFilePAGE;
            TeFormHttp.addEventListener("readystatechange", function () {
                switch (TeFormHttp.status) {
                    case 0:
                        break;
                    case 200:
                        if (TeFormHttp.readyState == 4) {
                            resolve(TeFormHttp.responseText);
                            TeFormButtonThis.disabled = false;
                            TeFormButtonThis.innerHTML = ButtonInnerHTML;
                        }
                        break;
                    default:
                        TeFormError += "-" + TeFormHttp.status + "-" + TeFormHttp.readyState;
                        reject("编号:" + TeFormError + "\n表单发送错误，请提供错误编号给客服人员。");
                        TeFormButtonThis.disabled = false;
                        TeFormButtonThis.innerHTML = ButtonInnerHTML;
                        break;
                }
            });
            TeFormHttp.open(TeFormDataMethod, TeFormUrl + TeFormDataURL);
            TeFormHttp.send(new FormData(newFormData));
        } else {
            reject("Error:404-n");
        }
    });
}