Imports DelSole.VSIX
Imports Microsoft.Win32
Imports Syncfusion.Windows.Edit
Imports <xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports Syncfusion.UI.Xaml.Grid
Imports System.ComponentModel
Imports DelSole.VSIX.VsiTools, DelSole.VSIX.SnippetTools
Imports System.Diagnostics
Imports EnvDTE
Imports System.Windows.Navigation
Imports System.Linq
Imports System
Imports System.Windows, System.Windows.Controls
Imports System.Xml.Linq
Imports System.Collections.Generic

'''<summary>
''' Interaction logic for CodeSnippetStudioToolWindowControl.xaml
'''</summary>
Partial Public Class CodeSnippetStudioToolWindowControl
    Inherits System.Windows.Controls.UserControl

    Private theData As VSIXPackage
    Private snippetProperties As CodeSnippet
    Private Property [Imports] As [Imports]
    Private Property References As References
    Private Property Declarations As Declarations

    Private Sub ResetPkg()
        Me.theData = New VSIXPackage

        Me.DataContext = Me.theData
        Me.PackageTab.Focus()
    End Sub

    Private Sub Hyperlink_RequestNavigate(ByVal sender As Object, ByVal e As RequestNavigateEventArgs)
        System.Diagnostics.Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
        e.Handled = True
    End Sub

    Private Sub MyControl_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        ResetPkg()
        Me.RootTabControl.SelectedIndex = 0
        Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.VisualBasic
        Me.LanguageCombo.SelectedIndex = 0
        Me.SnippetPropertyGrid.SelectedObject = New CodeSnippet
        'Me.SnippetPropertyGrid.DescriptionPanelVisibility = Visibility.Visible
        Me.Imports = New [Imports]
        Me.Declarations = New Declarations
        Me.References = New References

        Me.ImportsDataGrid.ItemsSource = Me.Imports
        Me.RefDataGrid.ItemsSource = Me.References
        Me.DeclarationsDataGrid.ItemsSource = Me.Declarations
    End Sub

    Private Sub AddSnippetsButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        With dlg
            .Multiselect = True
            .Filter = "Snippet files (*.snippet)|*.snippet|All files|*.*"

            .Title = "Select code snippets"
            If .ShowDialog = True Then

                For Each item In .FileNames
                    Dim sninfo As New SnippetInfo
                    sninfo.SnippetFileName = IO.Path.GetFileName(item)
                    sninfo.SnippetPath = IO.Path.GetDirectoryName(item)
                    sninfo.SnippetLanguage = SnippetInfo.GetSnippetLanguage(item)
                    sninfo.SnippetDescription = SnippetInfo.GetSnippetDescription(item)
                    Me.theData.CodeSnippets.Add(sninfo)
                Next

            Else
                Exit Sub
            End If
        End With
    End Sub

    Private Sub BuildVsixButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        If Me.theData.HasErrors Then
            System.Windows.MessageBox.Show("The metadata information is incomplete. Fix errors before compiling.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        If theData.CodeSnippets.Any = False Then
            System.Windows.MessageBox.Show("The code snippet list is empty. Please add at least one before proceding.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        Dim testLang = Me.theData.TestLanguageGroup()

        If testLang = False Then
            System.Windows.MessageBox.Show("You have added code snippets of different programming languages. " + Environment.NewLine + "VSIX packages offer the best customer experience possible with snippets of only one language." +
                                           "For this reason, leave snippets of only one language and remove others before building the package.", "Code Snippet Studio", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning)
            Exit Sub
        End If

        Try
            Dim dlg As New SaveFileDialog
            With dlg
                .Title = "Specify the .vsix name and location"
                .OverwritePrompt = True
                .Filter = "VSIX packages|*.vsix"
                If .ShowDialog = True Then
                    Me.theData.Build(.FileName)
                    Dim result = MessageBox.Show("Package " + IO.Path.GetFileName(.FileName) + " created. Would you like to install the package for testing now?", "Code Snippet Studio", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question)
                    If result = System.Windows.MessageBoxResult.No Then
                        Exit Sub
                    Else
                        Diagnostics.Process.Start(.FileName)
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            System.Windows.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub GuideButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles GuideButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.doc, *.docx, *.rtf, *.txt, *.htm, *.html)|*.doc;*.docx;*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.GettingStartedGuide = .FileName
            End If
        End With
    End Sub

    Private Sub IconButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles IconButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select icon file"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.IconPath = .FileName
            End If
        End With
    End Sub

    Private Sub ImageButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles ImageButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select preview image"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.PreviewImagePath = .FileName
            End If
        End With
    End Sub

    Private Sub LicenseButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles LicenseButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt)|*.rtf;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.License = .FileName
            End If
        End With
    End Sub

    Private Sub RelNotesButton_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles RelNotesButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt, *.htm, *.html)|*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.ReleaseNotes = .FileName
            End If
        End With
    End Sub

    Private Sub RemoveSnippetButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        Try

            For Each item In Me.CodeSnippetsDataGrid.SelectedItems.Cast(Of SnippetInfo).ToList
                Me.theData.CodeSnippets.Remove(TryCast(item, SnippetInfo))
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ResetPkgButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        ResetPkg()
    End Sub

    Private Sub AboutButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        System.Diagnostics.Process.Start("https://github.com/AlessandroDelSole/CodeSnippetStudio")
    End Sub

    Private Sub VsiButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)
        If theData.HasErrors Then
            MessageBox.Show("The package metadata contain errors that must be fixed before performing a conversion." & Environment.NewLine &
                            "Please go to the Package and Share tab and check values in the Metadata nested tab.")
            Exit Sub
        End If

        Dim dlg As New OpenFileDialog
        Dim inputFile As String, outputFile As String

        With dlg
            .Title = "Select .vsi file"
            .Filter = "Visual Studio Community Installer (*.vsi)|*.vsi|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            inputFile = .FileName
        End With

        Dim dlg2 As New SaveFileDialog
        With dlg2
            .OverwritePrompt = True
            .Title = "Output .vsix file"
            .Filter = ".vsix files (*.vsix)|*.vsix|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            outputFile = .FileName
        End With

        VsiService.Vsi2Vsix(inputFile, outputFile, theData.SnippetFolderName,
                             theData.PackageAuthor, theData.ProductName,
                             theData.PackageDescription,
                             theData.IconPath, theData.PreviewImagePath,
                             theData.MoreInfoURL)
        MessageBox.Show($"Successfully converted {inputFile} into {outputFile}")
    End Sub

    Private Sub SignButton_Click(sender As Object, e As RoutedEventArgs)
        If PfxPassword.Password.Length = 0 Then
            MessageBox.Show("Please enter the password for the certificate file.")
            Exit Sub
        End If

        If PfxTextBox.Text = "" Or String.IsNullOrEmpty(PfxTextBox.Text) Then
            MessageBox.Show("Please specify a valid X.509 certificate file.")
            Exit Sub
        End If

        If Not IO.File.Exists(PfxTextBox.Text) Then
            MessageBox.Show("The specified certificate file does not exist.")
            Exit Sub
        End If

        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select the .vsix you want to sign"
            .Filter = "VSIX packages (*.vsix)|*.pfx|All files|*.*"
            If .ShowDialog = True Then
                VSIXPackage.SignVsix(.FileName, PfxTextBox.Text, PfxPassword.Password)
                MessageBox.Show($"{ .FileName} signed successfully.")
            End If
        End With
    End Sub

    Private Sub PfxButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select certificate file"
            .Filter = "Certificate files (*.pfx)|*.pfx|All files|*.*"
            If .ShowDialog = True Then
                Me.PfxTextBox.Text = .FileName
            End If
        End With
    End Sub

    Private Sub LanguageCombo_SelectionChanged(sender As Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles LanguageCombo.SelectionChanged
        Dim cb = CType(sender, ComboBox)
        Select Case cb.SelectedIndex
            Case = 0
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.VisualBasic
                EnableDataGrids()
            Case = 1
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.CSharp
                DisableDataGrids()
            Case = 2
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.SQL
                DisableDataGrids()
            Case = 3
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML
                DisableDataGrids()
            Case = 4
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.Text
                DisableDataGrids()
            Case = 5
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML
                DisableDataGrids()
            Case = 6
                Me.editControl1.DocumentLanguage = Syncfusion.Windows.Edit.Languages.XML
                DisableDataGrids()
        End Select
    End Sub

    Private Sub SaveSnippetButton_Click(sender As Object, e As RoutedEventArgs)
        Me.snippetProperties = CType(Me.SnippetPropertyGrid.SelectedObject, CodeSnippet)

        If snippetProperties.Author = "" Or String.IsNullOrEmpty(snippetProperties.Author) Then
            MessageBox.Show("Snippet author is missing.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        If snippetProperties.Title = "" Or String.IsNullOrEmpty(snippetProperties.Title) Then
            MessageBox.Show("Snippet title is missing.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        If snippetProperties.Description = "" Or String.IsNullOrEmpty(snippetProperties.Description) Then
            MessageBox.Show("Snippet description is missing.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        Dim dlg2 As New SaveFileDialog
        With dlg2
            .OverwritePrompt = True
            .Title = "Output .snippet file"
            .Filter = ".snippet files (*.snippet)|*.snippet|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If

            SaveSnippet(.FileName)
        End With
    End Sub

    Private Sub SaveSnippet(fileName As String)
        Dim currentLang = CType(LanguageCombo.SelectedItem, ComboBoxItem).Tag.ToString()
        Dim selectedSnippet = CType(SnippetPropertyGrid.SelectedObject, CodeSnippet)

        Dim keywords As IEnumerable(Of String)
        If selectedSnippet.Keywords Is Nothing Then
            keywords = New List(Of String) From {String.Empty}
        Else
            keywords = selectedSnippet?.Keywords?.Split(","c).AsEnumerable
        End If

        SnippetService.SaveSnippet(fileName, selectedSnippet.Kind, currentLang, selectedSnippet.Title,
                               selectedSnippet.Description, selectedSnippet.HelpUrl,
                               selectedSnippet.Author, selectedSnippet.Shortcut, editControl1.Text,
                               Me.Imports, References, Declarations, keywords)
        MessageBox.Show($"{fileName} saved correctly.")
    End Sub

    Private Sub ExtractButton_Click(sender As Object, e As RoutedEventArgs)
        Dim dlg As New OpenFileDialog
        Dim inputFile As String, outputFolder As String

        With dlg
            .Title = "Select .vsix file"
            .Filter = "Visual Studio Extension (*.vsix)|*.vsix|All files|*.*"
            If Not .ShowDialog = True Then
                Exit Sub
            End If
            inputFile = .FileName
        End With

        Dim dlg2 As New CommonOpenFileDialog()
        dlg2.Title = "Destination folder"
        dlg2.IsFolderPicker = True
        dlg2.InitialDirectory = Environment.SpecialFolder.MyDocuments
        dlg2.DefaultDirectory = Environment.SpecialFolder.MyDocuments
        dlg2.EnsureFileExists = True
        dlg2.EnsurePathExists = True
        dlg2.EnsureValidNames = True
        dlg2.Multiselect = False
        dlg2.ShowPlacesList = True

        If Not dlg2.ShowDialog = CommonFileDialogResult.Ok Then
            Exit Sub
        End If

        outputFolder = dlg2.FileName

        VSIXPackage.ExtractVsix(inputFile, outputFolder, OnlySnippetsCheckBox.IsChecked)
        MessageBox.Show($"Successfully extracted {inputFile} into {outputFolder}")
    End Sub

    Private Sub DeclarationsDataGrid_AddNewRowInitiating(sender As Object, args As AddNewRowInitiatingEventArgs) Handles DeclarationsDataGrid.AddNewRowInitiating
        Dim item = CType(args.NewObject, Declaration)
        item.ID = editControl1.SelectedText
        item.Default = editControl1.SelectedText
    End Sub

    Private Sub DisableDataGrids()
        Me.ImportsDataGrid.IsEnabled = False
        Me.RefDataGrid.IsEnabled = False
    End Sub

    Private Sub EnableDataGrids()
        Me.ImportsDataGrid.IsEnabled = True
        Me.RefDataGrid.IsEnabled = True
    End Sub

    Private Shared Function ReturnSnippetKind(kind As SnippetTools.CodeSnippetKinds) As String
        Dim snippetKind As String
        Select Case kind
            Case SnippetTools.CodeSnippetKinds.MethodBody
                snippetKind = "method body"
            Case SnippetTools.CodeSnippetKinds.MethodDeclaration
                snippetKind = "method decl"
            Case SnippetTools.CodeSnippetKinds.File
                snippetKind = "file"
            Case SnippetTools.CodeSnippetKinds.TypeDeclaration
                snippetKind = "type decl"
            Case Else
                snippetKind = "any"
        End Select

        Return snippetKind
    End Function

    Public Shared Sub SaveSnippet1(fileName As String, kind As SnippetTools.CodeSnippetKinds,
                        snippetLanguage As String, snippetTitle As String,
                        snippetDescription As String, snippetHelpUrl As String,
                        snippetAuthor As String, snippetShortcut As String,
                        snippetCode As String, importDirectives As [Imports],
                        references As References, declarations As Declarations,
                        keywords As IEnumerable(Of String))

        Dim snippetKind As String = ReturnSnippetKind(kind)

        Dim editedCode = snippetCode
        For Each decl In declarations
            editedCode = editedCode.Replace(decl.Default, "$" & decl.ID & "$")
        Next

        If keywords Is Nothing Then
            keywords = New List(Of String) From {String.Empty}
        End If

        Dim cdata As New XCData(editedCode)
        Dim doc = <?xml version="1.0" encoding="utf-8"?>
                  <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
                      <CodeSnippet Format="1.0.0">
                          <Header>
                              <Title><%= snippetTitle %></Title>
                              <Author><%= snippetAuthor %></Author>
                              <Description><%= snippetDescription %></Description>
                              <HelpUrl><%= snippetHelpUrl %></HelpUrl>
                              <SnippetTypes>
                                  <SnippetType>Expansion</SnippetType>
                              </SnippetTypes>
                              <Keywords>
                                  <%= From key In keywords
                                      Select <Keyword><%= key %></Keyword> %>
                              </Keywords>
                              <Shortcut><%= snippetShortcut %></Shortcut>
                          </Header>
                          <Snippet>
                              <References>
                                  <%= From ref In references
                                      Select <Reference>
                                                 <Assembly><%= ref.Assembly %></Assembly>
                                                 <Url><%= ref.Url %></Url>
                                             </Reference> %>
                              </References>
                              <Imports>
                                  <%= From imp In importDirectives
                                      Select <Import>
                                                 <Namespace><%= imp.ImportDirective %></Namespace>
                                             </Import> %>
                              </Imports>
                              <Declarations>
                                  <%= From decl In declarations
                                      Where decl.ReplacementType.ToLower = "object"
                                      Select <Object Editable="true">
                                                 <ID><%= decl.ID %></ID>
                                                 <Type><%= decl.Type %></Type>
                                                 <ToolTip><%= decl.ToolTip %></ToolTip>
                                                 <Default><%= decl.Default %></Default>
                                                 <Function><%= decl.Function %></Function>
                                             </Object> %>
                                  <%= From decl In declarations
                                      Where decl.ReplacementType.ToLower = "literal"
                                      Select <Literal Editable="true">
                                                 <ID><%= decl.ID %></ID>
                                                 <ToolTip><%= decl.ToolTip %></ToolTip>
                                                 <Default><%= decl.Default %></Default>
                                                 <Function><%= decl.Function %></Function>
                                             </Literal> %>
                              </Declarations>
                              <Code Language=<%= snippetLanguage %> Kind=<%= snippetKind %>
                                  Delimiter="$"></Code>
                          </Snippet>
                      </CodeSnippet>
                  </CodeSnippets>

        doc...<Code>.First.Add(cdata)

        doc.Save(fileName)
    End Sub
End Class