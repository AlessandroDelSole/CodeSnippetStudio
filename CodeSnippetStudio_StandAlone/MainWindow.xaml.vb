Imports System.Collections.ObjectModel
Imports DelSole.VSIX
Imports Microsoft.Win32
Imports Syncfusion.Windows.Edit
Imports <xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
Imports Microsoft.WindowsAPICodePack.Dialogs
Imports Syncfusion.UI.Xaml.Grid

Class MainWindow
    Private theData As VSIXPackage
    Private snippetProperties As SnippetFile
    Private Property [Imports] As New [Imports]
    Private Property References As New References
    Private Property Declarations As New Declarations

    Private Sub ResetPkg()
        Me.theData = New VSIXPackage

        Me.DataContext = Me.theData
        Me.PackageTab.Focus()
    End Sub

    Private Sub MyControl_Loaded(sender As Object, e As Windows.RoutedEventArgs) Handles Me.Loaded
        ResetPkg()
        Me.RootTabControl.SelectedIndex = 0
        Me.editControl1.DocumentLanguage = Languages.VisualBasic
        Me.LanguageCombo.SelectedIndex = 0
        Me.SnippetPropertyGrid.DescriptionPanelVisibility = Visibility.Visible
        Me.ImportsDataGrid.ItemsSource = Me.Imports
        Me.RefDataGrid.ItemsSource = Me.References
        Me.DeclarationsDataGrid.ItemsSource = Me.Declarations
    End Sub


    Private Sub AddSnippetsButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Dim dlg As New Microsoft.Win32.OpenFileDialog
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

    Private Sub BuildVsixButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        If Me.theData.HasErrors Then
            System.Windows.MessageBox.Show("The metadata information is incomplete. Fix errors before compiling.", "Snippet Package Builder", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        If theData.CodeSnippets.Any = False Then
            MessageBox.Show("The code snippet list is empty. Please add at least one before proceding.", "Snippet Package Builder", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Error)
            Exit Sub
        End If

        Dim testLang = Me.theData.TestLanguageGroup()

        If testLang = False Then
            System.Windows.MessageBox.Show("You have added code snippets of different programming languages. " + Environment.NewLine + "VSIX packages offer the best customer experience possible with snippets of only one language." +
                                           "For this reason, leave snippets of only one language and remove others before building the package.", "Snippet Package Builder", Windows.MessageBoxButton.OK, Windows.MessageBoxImage.Warning)
            Exit Sub
        End If

        Try
            Dim dlg As New Microsoft.Win32.SaveFileDialog
            With dlg
                .Title = "Specify the .vsix name and location"
                .OverwritePrompt = True
                .Filter = "VSIX packages|*.vsix"
                If .ShowDialog = True Then
                    Me.theData.Build(.FileName)
                    Dim result = MessageBox.Show("Package " + IO.Path.GetFileName(.FileName) + " created. Would you like to install the package for testing now?", "Snippet Package Builder", Windows.MessageBoxButton.YesNo, Windows.MessageBoxImage.Question)
                    If result = Windows.MessageBoxResult.No Then
                        Exit Sub
                    Else
                        Process.Start(.FileName)
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            System.Windows.MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub GuideButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles GuideButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.doc, *.docx, *.rtf, *.txt, *.htm, *.html)|*.doc;*.docx;*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.GettingStartedGuide = .FileName
            End If
        End With
    End Sub

    Private Sub IconButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles IconButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select icon file"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.IconPath = .FileName
            End If
        End With
    End Sub

    Private Sub ImageButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles ImageButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select preview image"
            .Filter = "All supported files (*.jpg, *.png, *.ico, *.bmp, *.tif, *.tiff, *.gif)|*.jpg;*.png;*.ico;*.bmp;*.tiff;*.tif;*.gif|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.PreviewImagePath = .FileName
            End If
        End With
    End Sub

    Private Sub LicenseButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles LicenseButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt)|*.rtf;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.License = .FileName
            End If
        End With
    End Sub

    Private Sub RelNotesButton_Click(sender As Object, e As Windows.RoutedEventArgs) Handles RelNotesButton.Click
        Dim dlg As New OpenFileDialog
        With dlg
            .Title = "Select documentation file"
            .Filter = "All supported files (*.rtf, *.txt, *.htm, *.html)|*.rtf;*.htm;*.html;*.txt|All files|*.*"
            If .ShowDialog = True Then
                Me.theData.ReleaseNotes = .FileName
            End If
        End With
    End Sub

    Private Sub RemoveSnippetButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Try

            For Each item In Me.CodeSnippetsDataGrid.SelectedItems.Cast(Of SnippetInfo).ToList
                Me.theData.CodeSnippets.Remove(TryCast(item, SnippetInfo))
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ResetPkgButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        ResetPkg()
    End Sub

    Private Sub AboutButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        Process.Start("http://visualstudiogallery.msdn.microsoft.com/de44d368-bab1-43cb-9167-701ac668a09c?SRC=Home")
    End Sub

    Private Sub VsiButton_Click(sender As Object, e As Windows.RoutedEventArgs)
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

        VSIXPackage.Vsi2Vsix(inputFile, outputFile, theData.SnippetFolderName, theData.PackageAuthor, theData.ProductName, theData.PackageDescription, theData.IconPath, theData.PreviewImagePath, theData.MoreInfoURL)
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
                Me.editControl1.DocumentLanguage = Languages.VisualBasic
            Case = 1
                Me.editControl1.DocumentLanguage = Languages.CSharp
            Case = 2
                Me.editControl1.DocumentLanguage = Languages.SQL
            Case = 3
                Me.editControl1.DocumentLanguage = Languages.XML
        End Select
    End Sub

    Private Sub SaveSnippetButton_Click(sender As Object, e As RoutedEventArgs)
        Me.snippetProperties = CType(Me.SnippetPropertyGrid.SelectedObject, SnippetFile)

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

        Dim cdata As New XCData(editControl1.Text)

        Dim doc = <?xml version="1.0" encoding="utf-8"?>
                  <CodeSnippets xmlns="http://schemas.microsoft.com/VisualStudio/2005/CodeSnippet">
                      <CodeSnippet Format="1.0.0">
                          <Header>
                              <Title><%= snippetProperties.Title %></Title>
                              <Author><%= snippetProperties.Author %></Author>
                              <Description><%= snippetProperties.Description %></Description>
                              <HelpUrl><%= snippetProperties.HelpUrl %></HelpUrl>
                              <SnippetTypes>
                                  <SnippetType>Expansion</SnippetType>
                              </SnippetTypes>
                              <Keywords>
                                  <Keyword></Keyword>
                              </Keywords>
                              <Shortcut><%= snippetProperties.Shortcut %></Shortcut>
                          </Header>
                          <Snippet>
                              <References>
                                  <%= From ref In Me.References
                                      Select <Reference>
                                                 <Assembly><%= ref.Assembly %></Assembly>
                                                 <Url><%= ref.Url %></Url>
                                             </Reference> %>
                              </References>
                              <Imports>
                                  <%= From imp In Me.Imports
                                      Select <Import>
                                                 <Namespace><%= imp.ImportDirective %></Namespace>
                                             </Import> %>
                              </Imports>
                              <Declarations>
                                  <%= From decl In Me.Declarations
                                      Select <Object Editable="true">
                                                 <ID><%= decl.ID %></ID>
                                                 <Type><%= decl.Type %></Type>
                                                 <ToolTip><%= decl.ToolTip %></ToolTip>
                                                 <Default><%= decl.Default %></Default>
                                                 <Function><%= decl.Function %></Function>
                                             </Object> %>
                              </Declarations>
                              <Code Language=<%= currentLang %> Kind="" Delimiter="$"></Code>
                          </Snippet>
                      </CodeSnippet>
                  </CodeSnippets>

        doc...<Code>.First.Add(cdata)

        doc.Save(fileName)
        MessageBox.Show($"{fileName} saved correctly")
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

        VSIXPackage.ExtractVsix(inputFile, outputFolder)
        MessageBox.Show($"Successfully extracted {inputFile} into {outputFolder}")
    End Sub

    Private Sub DeclarationsDataGrid_AddNewRowInitiating(sender As Object, args As AddNewRowInitiatingEventArgs) Handles DeclarationsDataGrid.AddNewRowInitiating
        Dim item = CType(args.NewObject, CodeObject)
        item.ID = editControl1.SelectedText
        item.Default = editControl1.SelectedText
    End Sub
End Class

Class Import
    Property ImportDirective As String
End Class

Class [Imports]
    Inherits ObservableCollection(Of Import)
End Class

Class Reference
    Property Assembly As String
    Property Url As String
End Class

Class References
    Inherits ObservableCollection(Of Reference)
End Class

Class CodeObject
    Property Editable As Boolean = True
    Property ID As String
    Property [Type] As String
    Property ToolTip As String
    Property [Default] As String
    Property [Function] As String
End Class

Class Declarations
    Inherits ObservableCollection(Of CodeObject)
End Class