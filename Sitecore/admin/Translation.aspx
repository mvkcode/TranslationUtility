<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Translation.aspx.cs" Inherits="UtilityProject.Translation" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Translation</title>
</head>
<body>
    <form id="form1" runat="server">
        <h1>Translation Utility</h1>
        <br />
        <br />
        <div><b>STEP 1 : </b>Generate XML from Sitecore. Click - <a href="/sitecore/shell/default.aspx?xmlcontrol=ExportLanguage" target="_blank">Export Languages</a> </div>
        <%--<div><b>STEP 1 : </b>Generate XML via Sitecore Control Panel - <a href="/sitecore/client/Applications/ControlPanel.aspx?sc_bw=1" target="_blank">Click Export Languages</a> </div>--%>
        <br />
        <br />
        <div>
            <b>STEP 2 : </b>Choose XML file &nbsp;<asp:FileUpload ID="FileUpload1" runat="server" /> 
        </div>
        <br />
        

         <div style="margin:10px 40px 10px 40px" >
            Click on <b>Download Excel</b> button to convert XML into Excel. 
            <asp:Button ID="btnDownload" runat="server" Text="Download Excel" OnClick="btnDownload_Click" />

            <%--<b>STEP 2 : </b>Choose file for translation &nbsp;
            <asp:Button ID="btnDownload" runat="server" Text="Download Content" OnClick="btnDownload_Click" />--%>
        </div>
        
        <br />
        <br />
        <div>
            <b>STEP 3 : </b>Upload translated Excel file  &nbsp;<asp:FileUpload ID="fuTranslator" runat="server" />
        </div>
        <br />
        
        <div style="margin:10px 40px 10px 40px">
            Click on <b>Update Translation XML</b> button to update XML with translated Excel. 
            <asp:Button ID="btnUpdateTranslation" runat="server" Text="Update Translation XML" OnClick="btnUpdateTranslation_Click" />
        </div>

        <br />
        <br />
        <%--<div><b>STEP 4 : </b>Update XML via Sitecore Control Panel - <a href="/sitecore/client/Applications/ControlPanel.aspx?sc_bw=1" target="_blank">Click Import Languages</a> </div>--%>

        <div><b>STEP 4 : </b>Update translated content into Sitecore. Click - <a href="/sitecore/shell/default.aspx?xmlcontrol=ImportLanguage" target="_blank">Import Languages</a> </div>


        <br />
        <br />
        <div>
            <asp:Label ID="lblMsg" runat="server" Text=""></asp:Label>
        </div>
    </form>
</body>
</html>
