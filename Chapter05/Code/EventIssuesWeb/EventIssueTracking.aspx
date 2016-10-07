<%@ Page Language="vb" AutoEventWireup="false" Codebehind="EventIssueTracking.aspx.vb" Inherits="EventIssues.EventIssueTracking"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
    <title>EventIssueTracking</title>
    <meta name="vs_showGrid" content="True">
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
  </HEAD>
  <body MS_POSITIONING="GridLayout">
    <form id="Form1" method="post" runat="server">
      <asp:Label id="lblHeading" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 16px"
        runat="server" Font-Names="Arial" Font-Size="X-Large">Service Failure Status</asp:Label>
      <asp:Label id="lbl_CurrentStatus_Value" style="Z-INDEX: 112; LEFT: 144px; POSITION: absolute; TOP: 264px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Width="464px" BorderColor="Black" BorderStyle="Solid"
        BorderWidth="1px"></asp:Label>
      <asp:Label id="lbl_CurrentStatus_Label" style="Z-INDEX: 111; LEFT: 16px; POSITION: absolute; TOP: 264px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Font-Bold="True" BackColor="#E0E0E0" Width="120px"
        BorderColor="Black" BorderStyle="Solid" BorderWidth="1px">Current Status</asp:Label>
      <asp:Label id="lbl_IssueDescription_Value" style="Z-INDEX: 108; LEFT: 144px; OVERFLOW: auto; POSITION: absolute; TOP: 120px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Height="136px" Width="464px" BorderColor="Black"
        BorderStyle="Solid" BorderWidth="1px"></asp:Label>
      <asp:Label id="lbl_CustomerEmail_Value" style="Z-INDEX: 107; LEFT: 144px; POSITION: absolute; TOP: 96px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Width="464px" BorderColor="Black" BorderStyle="Solid"
        BorderWidth="1px"></asp:Label>
      <asp:Label id="lbl_CustomerName_Value" style="Z-INDEX: 106; LEFT: 144px; POSITION: absolute; TOP: 72px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Width="464px" BorderStyle="Solid" BorderWidth="1px"></asp:Label>
      <asp:Label id="lbl_IssueDescription_Label" style="Z-INDEX: 105; LEFT: 16px; POSITION: absolute; TOP: 120px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Font-Bold="True" BackColor="#E0E0E0" Width="120px"
        BorderColor="Black" BorderStyle="Solid" BorderWidth="1px">Issue Description</asp:Label>
      <asp:Label id="lbl_CustomerEmail_Label" style="Z-INDEX: 104; LEFT: 16px; POSITION: absolute; TOP: 96px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Font-Bold="True" BackColor="#E0E0E0" Width="120px"
        BorderColor="Black" BorderStyle="Solid" BorderWidth="1px">Email Address</asp:Label>
      <HR style="Z-INDEX: 102; LEFT: 16px; WIDTH: 600px; POSITION: absolute; TOP: 56px" width="86.9%"
        SIZE="3" color="black">
      <asp:Label id="lbl_CustomerName_Label" style="Z-INDEX: 103; LEFT: 16px; POSITION: absolute; TOP: 72px"
        runat="server" Font-Names="Arial" Font-Size="10pt" Font-Bold="True" BackColor="#E0E0E0" Width="120px"
        BorderColor="Black" BorderStyle="Solid" BorderWidth="1px">Customer Name</asp:Label>
      <asp:Label id="lblTasks" style="Z-INDEX: 109; LEFT: 16px; POSITION: absolute; TOP: 320px" runat="server"
        Font-Names="Arial" Font-Size="Medium">Tasks Required to Resolve Issue</asp:Label>
      <asp:Panel id="panel_RepHolder" style="Z-INDEX: 110; LEFT: 16px; POSITION: absolute; TOP: 416px"
        runat="server" Height="352px">
        <asp:Repeater id="repTasks" runat="server">
          <ItemTemplate>
            <table style="width: 590px; border: 1px solid black;" cellspacing="0" cellpadding="3">
              <tr style="font-family: arial; font-size:10pt;">
                <td style="width:125px;"><b>Task Description:</b></td>
                <TD>
                  <asp:Label ID="lblTask" Runat="server"></asp:Label></TD>
              </tr>
              <tr style="font-family: arial; font-size:10pt; background-color: #DDDDDD;">
                <td style="border-top: 1px solid black;"><B>Task Status:</B></td>
                <td style="border-top: 1px solid black;">
                  <asp:Label ID="lblTaskStatus" Runat="server"></asp:Label>&nbsp;</td>
            </table>
            <br>
          </ItemTemplate>
        </asp:Repeater>
      </asp:Panel>
      <HR style="Z-INDEX: 113; LEFT: 16px; WIDTH: 600px; POSITION: absolute; TOP: 344px" width="86.9%"
        color="black" SIZE="3">
      <asp:Label id="lblDescription" style="Z-INDEX: 114; LEFT: 16px; POSITION: absolute; TOP: 352px"
        runat="server" Font-Size="10pt" Font-Names="Arial" Width="584px" Height="30px">The following tasks need to be carried out in order to resolve the aforementioned issue.  You can view the current status of each task by looking at the "Task Status" located after each Task listing.  When all tasks have been completed, the issue will be resolved.</asp:Label>
    </form>
  </body>
</HTML>
