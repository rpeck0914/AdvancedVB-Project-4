<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MasterPage.Master" CodeBehind="Suppliers.aspx.vb" Inherits=".WebForm4" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <link href="SuppliersStyleSheet.css" rel="stylesheet" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <h2>Suppliers</h2>
    <p></p>
    <asp:GridView ID="GridView1" runat="server"></asp:GridView>
</asp:Content>
