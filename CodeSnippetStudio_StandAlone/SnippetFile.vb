Imports System.ComponentModel

Public Class SnippetFile

    <Category("Identity")>
    <DisplayName("Author")>
    <Description("Author of code snippet. This is a required value.")>
    Public Property Author As String

    <Category("Identity")>
    <DisplayName("Title")>
    <Description("Title of code snippet. This is a required value.")>
    Public Property Title As String

    <Category("Identity")>
    <DisplayName("Description")>
    <Description("Description of code snippet. This is a required value.")>
    Public Property Description As String

    <Category("Identity")>
    <DisplayName("HelpUrl")>
    <Description("URL where users can find help. This is an optional value.")>
    Public Property HelpUrl As String

    <Category("Identity")>
    <DisplayName("Shortcut")>
    <Description("Keyboard shortcut for IntelliSense. This is an optional value.")>
    Public Property Shortcut As String
End Class
