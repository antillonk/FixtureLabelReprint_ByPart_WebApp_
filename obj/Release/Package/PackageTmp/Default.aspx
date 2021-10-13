<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="FixtureLabelReprint_ByPart_WebApp._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <br />
    <br />
    <asp:Label ID="lblJobNumText" runat="server" Text="Fixture Part #: "></asp:Label>
    <asp:TextBox ID="tbPartNum" runat="server"></asp:TextBox>&nbsp&nbsp
    # of Labels: 
    <asp:TextBox ID="tbLabelCount" runat="server" Width="41px" Text="1"></asp:TextBox>
    <br />
    <br />
   <%-- <asp:Label ID="lblQuantityText" runat="server" Text="Quantity: "></asp:Label>
    <asp:TextBox ID="tbQuantity" runat="server"></asp:TextBox>
    <br />
    <br />--%>
    <asp:Button ID="btnPrintLabels" runat="server" Text="Print Label" OnClick="btnPrintLabels_Click"/>
    <asp:Label ID="lblError" runat="server" Text="" Font-Italic="true" ForeColor="Red"></asp:Label>
</asp:Content>
