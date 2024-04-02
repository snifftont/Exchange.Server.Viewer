<%@ Page Title="" Language="vb" ValidateRequest="false"  AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="sendmessage.aspx.vb" Inherits="mailexchange.sendmessage" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
<script type="text/javascript" src="/Scripts/jscripts/tiny_mce/tiny_mce.js"></script>
<script type="text/javascript">
    tinyMCE.init({
        // General options
        mode: "textareas",
        theme: "advanced",
        encoding: "xml",
        plugins: "autolink,lists,spellchecker,pagebreak,style,layer,table,save,advhr,advimage,advlink,emotions,iespell,inlinepopups,insertdatetime,preview,media,searchreplace,print,contextmenu,paste,directionality,fullscreen,noneditable,visualchars,nonbreaking,xhtmlxtras,template",

        // Theme options
        theme_advanced_buttons1: "save,newdocument,|,bold,italic,underline,strikethrough,|,justifyleft,justifycenter,justifyright,justifyfull,|,styleselect,formatselect,fontselect,fontsizeselect",
        theme_advanced_buttons2: "cut,copy,paste,pastetext,pasteword,|,search,replace,|,bullist,numlist,|,outdent,indent,blockquote,|,undo,redo,|,link,unlink,anchor,image,cleanup,help,code,|,insertdate,inserttime,preview,|,forecolor,backcolor",
        theme_advanced_buttons3: "tablecontrols,|,hr,removeformat,visualaid,|,sub,sup,|,charmap,emotions,iespell,media,advhr,|,print,|,ltr,rtl,|,fullscreen",
        theme_advanced_buttons4: "insertlayer,moveforward,movebackward,absolute,|,styleprops,spellchecker,|,cite,abbr,acronym,del,ins,attribs,|,visualchars,nonbreaking,template,blockquote,pagebreak,|,insertfile,insertimage",
        theme_advanced_toolbar_location: "top",
        theme_advanced_toolbar_align: "left",
        theme_advanced_statusbar_location: "bottom",
        theme_advanced_resizing: true,

        // Skin options
        skin: "o2k7",
        skin_variant: "silver",

        // Example content CSS (should be your site CSS)
        content_css: "css/example.css",

        // Drop lists for link/image/media/template dialogs
        template_external_list_url: "js/template_list.js",
        external_link_list_url: "js/link_list.js",
        external_image_list_url: "js/image_list.js",
        media_external_list_url: "js/media_list.js",

        // Replace values for the template plugin
        template_replace_values: {
            username: "Some User",
            staffid: "991234"
        }
    });
</script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<div style="background-color:InfoBackground;">
                 <fieldset>
                 <legend>Send Message</legend>
                 <p>
                     <asp:Label ID="lblTo" runat="server" Text="To"></asp:Label>
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <asp:TextBox ID="txtTo" CssClass="textEntry" runat="server"></asp:TextBox>
                 </p>
                 <p>
                     <asp:Label ID="lblFrom" runat="server" Text="From"></asp:Label>
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:TextBox ID="txtFrom" CssClass="textEntry" runat="server"></asp:TextBox>
                 </p>
                 <p>
                     <asp:Label ID="lblSubject" runat="server" Text="Subject"></asp:Label>
                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                     <asp:TextBox ID="txtSubject" CssClass="textEntry" runat="server"></asp:TextBox>
                 </p>
                 <p>
                  <asp:Label ID="Label1" runat="server" Text="Body"></asp:Label>
                    &nbsp;:</p>
                    <p>
                     <asp:TextBox ID="txtBody" runat="server" TextMode="MultiLine" Height="500px" 
                         Width="396px"></asp:TextBox>
                 </p>
                 <div class="submitButtonsimple">
   <asp:Button ID="btnSend" runat="server" Text="Send" OnClick="btnSend_Click" />
</div>

</fieldset>
</div>
</asp:Content>
