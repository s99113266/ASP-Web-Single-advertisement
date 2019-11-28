<%@ Page Language="VBScript" AutoEventWireup="true" aspcompat=true Debug="true"%>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web.Security.FormsAuthentication"%>
<!--#include file="~/OledbConnection.aspx"-->

<%

    Dim z0f01_Guid As Guid = Guid.NewGuid()  '''亂數雜湊
    Dim z0f01_z0f1,z0f01_z0f2,z0f01_z0f3,z0f01_z0f4
    z0f01_z0f1 = Trim(Request.Form("z0f1"))
    z0f01_z0f2 = Trim(Request.Form("z0f2"))
    z0f01_z0f3 = Trim(Request.Form("z0f3"))
    z0f01_z0f4 = Request.Files("z0f4")


    
    Response.Write(tempCartId)

    con.open(ConnectionText(ConnectionDbFile, ConnectionDbPsw))



    con.close()
%>