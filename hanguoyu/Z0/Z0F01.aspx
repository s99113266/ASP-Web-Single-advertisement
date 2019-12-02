<%@ Page Language="VBScript" AutoEventWireup="true" aspcompat=true Debug="true"%>
<%@ Import Namespace="System"%>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web.Security.FormsAuthentication"%>
<!--#include file="~/Functions/TE_Bytes.inc"-->
<!--#include file="~/Functions/TE_TextReplace.inc"-->
<!--#include file="~/Functions/TE_Replace_BR_Tag.inc"-->
<!--#include file="~/Functions/TE_Re0_9.inc"-->
<!--#include file="~/OledbConnection.aspx"-->

<%

    Dim z0f01_Guid As Guid = Guid.NewGuid()  '''亂數雜湊
    Dim z0f01_z0f1,z0f01_z0f2,z0f01_z0f3,z0f01_z0f4
    Dim z0f01_Number    '''商品編號
    Dim z0f01_File, z0f01_FileImgName  '''圖片路徑, 圖片名稱
    Dim z0f01_Imagelist '''多選圖片陣列

    z0f01_Number = "C" & CStr(Format(now(),"yyMMddHHmmssffff"))
    z0f01_File = CStr(System.Web.HttpContext.Current.Request.PhysicalApplicationPath) & "mdb\CommodityImage\"

    z0f01_z0f1 = TETextReplace(Trim(Request.Form("z0f1")))
    if StrByteLen(z0f01_z0f1) < 1 Or StrByteLen(z0f01_z0f1) > 100 then
        Response.Write("商品名稱不能空白，或超過50個中文字。")
        return
    end if

    z0f01_z0f2 = Trim(Request.Form("z0f2"))
    if StrByteLen(z0f01_z0f2) < 1 then
        Response.Write("商品金額必須填寫，免費請填寫 0")
        return
    end if

    z0f01_z0f3 = Request.Files
    try
        if z0f01_z0f3.count > 0 then
            z0f01_Imagelist = ""
            z0f01_FileImgName = ""
            for i = 0 to z0f01_z0f3.count - 1
                z0f01_FileImgName = z0f01_Number & "(" & i & ")" & LCase(system.IO.Path.GetExtension(z0f01_z0f3(i).fileName))
                z0f01_z0f3(i).SaveAS(z0f01_File & z0f01_FileImgName)
                if Len(z0f01_Imagelist) = 0 then
                    z0f01_Imagelist = z0f01_FileImgName
                elseif Len(z0f01_Imagelist) > 0 then
                    z0f01_Imagelist &= "," & z0f01_FileImgName
                end if
            next
        end if
    catch err1 as exception
        Response.Write(err1.message & "{z0f01:err1-1}")
        return
    end try

    z0f01_z0f4 = TETextReplace(Trim(Request.Form("z0f4")))
    z0f01_z0f4 = TEReplace_VbcrlfToBR(z0f01_z0f4)

    Dim z0f01_CommText, z0f01_Execute
    try
        z0f01_Execute = ""
        z0f01_Execute &= "'" & z0f01_Number & "'"
        z0f01_Execute &= ",'" & z0f01_z0f1 & "'"
        z0f01_Execute &= "," & z0f01_z0f2
        z0f01_Execute &= ",'" & z0f01_Imagelist & "'"
        z0f01_Execute &= ",'" & z0f01_Guid.ToString("N") & "'"
        z0f01_Execute &= ",'" & z0f01_z0f4 & "'"
        con.open(ConnectionText(ConnectionDbFile, ConnectionDbPsw))
        z0f01_CommText = "Insert Into Commodity (Cod01, Cod02, Cod03, Cod04, Cod06, Cod09) Values (" & z0f01_Execute & ")"
        con.Execute(z0f01_CommText)
        con.close()
        Response.Write("1")
    catch err2 as exception
        Response.Write(err2.message & "{z0f01:err2-1}")
        return
    end try
%>