<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="emailviewer.aspx.vb" Inherits="mailexchange.emailviewer" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>    
    <asp:Panel ID="pnlpop1" CssClass="centerpopupmail" Visible="false" runat="server">
   
     <asp:Label ID="lblpwdmsg" runat="server" Text=""></asp:Label>
  
            
                 <fieldset>
                 <legend>Mail Content</legend>
                   <div class="submitButtonsimple" style="width:200px; float:right; text-align:right;">
   <asp:Button ID="btnClose3" runat="server" Text="Close" OnClick="btnClose3_Click" />
</div>
<p style="text-align:left;">
    <asp:Label ID="lblContent" runat="server" Text=""></asp:Label>
    
    </p>
   
    <center>
<div class="submitButtonsimple">
   <asp:Button ID="btnClose" runat="server" Text="Close" OnClick="btnClose_Click" />
</div>
</center>
</fieldset>
</asp:Panel>
<asp:Panel ID="pnlpop2" CssClass="centerpopupmail" Visible="false" runat="server">

     <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
  
            
                 <fieldset>
                 <legend>Mail Content</legend>
                  <div class="submitButtonsimple" style="width:100px; float:right; text-align:right;">
   <asp:Button ID="btnClose4" runat="server" Text="Close" OnClick="btnClose4_Click" />
</div>
                 <p><asp:HyperLink ID="hlnkReply" runat="server">Reply</asp:HyperLink>&nbsp;&nbsp;&nbsp;
                     <asp:HyperLink ID="hlnkForward" runat="server">Forward</asp:HyperLink>
                     
                     </p>
<p style="text-align:left;">
    <asp:Label ID="lblicontent" runat="server" Text=""></asp:Label>
    
    </p>

    <center>
<div class="submitButtonsimple">
   <asp:Button ID="btnClose2" runat="server" Text="Close" OnClick="btnClose2_Click" />
</div>
</center>
</fieldset>
</asp:Panel>

    <div>
    <p>
        Click here to <a href="/Account/changepassword.aspx">Change Password </a>.
    </p>
    <asp:DropDownList ID="ddlEmailof" runat="server" Width="150px" AutoPostBack="true" OnSelectedIndexChanged="ddlEmailof_SelectedIndexChanged">
    <asp:ListItem>Inbox</asp:ListItem>
    <asp:ListItem>Spam</asp:ListItem>
    </asp:DropDownList>
    
    </div>
    <center>
    <asp:UpdateProgress ID="UpdateProgress1" runat="server">
            <ProgressTemplate>
          
                <asp:Image ID="Image1" CssClass="progressbar" runat="server" ImageUrl="/Styles/ajax-progress.gif" />
          
            </ProgressTemplate>
        
            </asp:UpdateProgress>
            </center>
    <div>
     <fieldset>
                    <legend><asp:LinkButton ID="lnlEmailChecknew" runat="server">Check for New Messages</asp:LinkButton>
                   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:LinkButton ID="LinkButton1" runat="server">Load More Messages</asp:LinkButton>
                    </legend>
         Show<asp:DropDownList ID="ddlPage" runat="server" Height="16px" 
             AutoPostBack="True">
        <asp:ListItem>20</asp:ListItem>
    <asp:ListItem>50</asp:ListItem>
    <asp:ListItem>100</asp:ListItem>
    </asp:DropDownList>
         &nbsp;Per Page&nbsp;&nbsp;Total:&nbsp;<asp:Label ID="lblTotal" runat="server" Text=""></asp:Label>
        <br />
         <asp:Panel ID="Panel1" runat="server">

    
    
        <asp:GridView ID="GridView1" runat="server" Width="810px" 
                 AutoGenerateColumns="False" AllowPaging="True">
        <Columns>
     <%--   <asp:TemplateField>
     <ItemTemplate>
      
    <asp:CheckBox ID="CheckBox3" runat="server" />
     </ItemTemplate>  
      </asp:TemplateField> --%>
      <asp:TemplateField HeaderText="#">  
     <ItemTemplate>
     <asp:Label ID="lblread" runat="server" Visible="false" Text='<%# Eval("is_read") %>'></asp:Label>
     <asp:Label ID="lblSrNo" runat="server" Text='<%# Eval("SrNo") %>'></asp:Label>
     </ItemTemplate>
          <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
          <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
     </asp:TemplateField> 
     <asp:TemplateField HeaderText="Sender"> 
     <ItemTemplate>
      <asp:Label ID="lblFrom" runat="server" Text='<%# Eval("er_from") %>'></asp:Label>
     </ItemTemplate>
         <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
      </asp:TemplateField> 
     <asp:TemplateField HeaderText="Subject"> 
     <ItemTemplate>
      <asp:LinkButton ID="lnkSubject1" CommandName="mopen" CommandArgument='<%# Eval("er_id").ToString() %>' runat="server" Text='<%# Eval("er_subject") %>'></asp:LinkButton>
     
    <%-- <a style="text-decoration:none;" target="_blank"  href='<%# Eval("email_link") %>'>
     <asp:Label ID="lblSubject" runat="server" Text='<%# Eval("er_subject") %>'></asp:Label></a>--%>
     </ItemTemplate>
         <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
      </asp:TemplateField> 
     <asp:TemplateField HeaderText="Date"> 
     <ItemTemplate>
     <asp:Label ID="lblDate" runat="server" Text='<%# Eval("er_time") %>'></asp:Label>
     </ItemTemplate>
         <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
      </asp:TemplateField> 
     <asp:TemplateField HeaderText="Size"> 
     <ItemTemplate>
         <asp:Label ID="lblSize" runat="server" Text='<%# Eval("er_size") %>'></asp:Label>
     </ItemTemplate>
         <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:TemplateField> 
      <%-- <asp:TemplateField HeaderText="Forward"> 
     <ItemTemplate>
     <a style="text-decoration:none;" target="_blank"  href='<%# Eval("email_link") + "?cmd=forward" %>'>
         <asp:Label ID="lblForward" runat="server" Text="Forward"></asp:Label>
         </a>
     </ItemTemplate>
         <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:TemplateField>--%>
     <%-- <asp:TemplateField HeaderText="Reply"> 
     <ItemTemplate>
     <a style="text-decoration:none;" target="_blank"  href='<%# Eval("email_link") + "?cmd=reply" %>'>
         <asp:Label ID="lblReply" runat="server" Text="Reply"></asp:Label>
         </a>
     </ItemTemplate>
         <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:TemplateField>--%>
        </Columns>
        </asp:GridView>
 
 
    </asp:Panel>
     <asp:Panel ID="Panel2" runat="server">
     <div class="submitButtonsimple">
     <br />
       <asp:Button ID="btnDelSelected" runat="server" Text="Delete Selected" />&nbsp;&nbsp;
        <asp:Button ID="btnDelAll" runat="server" Text="Delete All" />&nbsp;&nbsp;
        <asp:Button ID="btndelivers" runat="server" Text="Deliver Selected" />
        <%-- <asp:LinkButton ID="LinkButton1" runat="server">Deliver Selected</asp:LinkButton>--%></div>
 <asp:GridView ID="GridView2" runat="server" Width="810px" AutoGenerateColumns="False" 
             AllowPaging="True">
        <Columns>
        <asp:TemplateField>
    
      
    <HeaderTemplate>
                    <asp:CheckBox runat="server" ID="chkall" AutoPostBack="True" OnCheckedChanged="chkall_CheckedChanged" />
                </HeaderTemplate>
                <ItemTemplate>
                 <asp:Label ID="lblread" runat="server" Visible="false" Text='<%# Eval("is_read") %>'></asp:Label>
                    <asp:Label ID="lblid" runat="server" Visible="false" Text='<%# Eval("er_id") %>'></asp:Label>
                    <asp:CheckBox runat="server" ID="chkSelect"  AutoPostBack="True" OnCheckedChanged="chkSelect_CheckedChanged" />
                </ItemTemplate>
     
      </asp:TemplateField> 
      <asp:TemplateField HeaderText="#">  
     <ItemTemplate>
     <asp:Label ID="lblSrNo" runat="server" Text='<%# Eval("SrNo") %>'></asp:Label>
     </ItemTemplate>
          <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
          <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
     </asp:TemplateField> 
     <asp:TemplateField HeaderText="Sender"> 
     <ItemTemplate>
      <asp:Label ID="lblFrom" runat="server" Text='<%# Eval("er_from") %>'></asp:Label>
     </ItemTemplate>
         <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
      </asp:TemplateField> 
     <asp:TemplateField HeaderText="Subject"> 
     <ItemTemplate>
         <asp:LinkButton ID="lnkSubject" CommandName="mopn" CommandArgument='<%# Eval("er_id").ToString() %>' runat="server" Text='<%# Eval("er_subject") %>'></asp:LinkButton>
     
     </ItemTemplate>
         <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
      </asp:TemplateField> 
     <asp:TemplateField HeaderText="Date"> 
     <ItemTemplate>
     <asp:Label ID="lblDate" runat="server" Text='<%# Eval("er_time") %>'></asp:Label>
     </ItemTemplate>
         <HeaderStyle Font-Bold="True" HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
      </asp:TemplateField> 
     <asp:TemplateField HeaderText="Size"> 
     <ItemTemplate>
         <asp:Label ID="lblSize" runat="server" Text='<%# Eval("er_size") %>'></asp:Label>
     </ItemTemplate>
         <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:TemplateField> 
       <asp:TemplateField HeaderText="Details"> 
     <ItemTemplate>
      <div class="submitButtonsimple">
          <asp:Button ID="btnContent" runat="server" Text="ViewContent" CommandName="viewcont" CommandArgument='<%# Eval("er_id").ToString() %>' />
      </div>
     </ItemTemplate>
         <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:TemplateField> 
       <asp:TemplateField HeaderText="Deliver"> 
     <ItemTemplate>
      <div class="submitButtonsimple">
          <asp:Button ID="btnDeliver" runat="server" Text="Deliver" CommandName="deliv" CommandArgument='<%# Eval("er_id").ToString() %>' />
      </div>
     </ItemTemplate>
         <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
         <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:TemplateField> 
        </Columns>
        </asp:GridView>
    </asp:Panel>
    <asp:LinkButton ID="lnkAllload" runat="server">Load More Messages</asp:LinkButton>
   </fieldset>
  
        

    </div>
    <div>
    <div style="float:left; text-align:left; width:48%">
     <fieldset>
                    <legend>Search</legend>
                 <div style="width:80%; float:left;">  
        <asp:RadioButton ID="rbtnSender" GroupName="email" runat="server" Text=" Sender" 
                         AutoPostBack="True" />&nbsp;<asp:RadioButton ID="rbtnSubject"
            runat="server" GroupName="email" Text=" Subject" AutoPostBack="True" />&nbsp; 
                     <asp:TextBox ID="txtSearch" runat="server"></asp:TextBox></div>
            <div class="submitButtonsimple"  style="width:20%; float:right;"> <asp:Button ID="btnGoSearch" runat="server" Text="Go" /></div>
           </fieldset>
    </div>
    <div style="float:right; text-align:right; width:40%;">
    <div class="submitButtonsimple">
   Show&nbsp;<asp:DropDownList ID="ddlFilter" runat="server" AutoPostBack="True">
    <asp:ListItem>All</asp:ListItem>
    <asp:ListItem>Last Week</asp:ListItem>
    <asp:ListItem>Last Month</asp:ListItem>
    </asp:DropDownList>
    
        <asp:Button ID="btnGoFilter" runat="server" Text="Go" /></div>
          
           </div>
           </div>
           </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
