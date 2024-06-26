Using ResourceX

Partial Class Form1
    Inherits System.Windows.Forms.Form

    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmbCulture = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        ' Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        ' Button1
        '
        Me.Button1.Location = New System.Drawing.Point(13, 40)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        '
        ' cmbCulture
        '
        Me.cmbCulture.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCulture.FormattingEnabled = True
        Me.cmbCulture.Items.AddRange(New Object() {"en-US", "en-CA", "fr-FR", "es-ES"})
        Me.cmbCulture.Location = New System.Drawing.Point(13, 70)
        Me.cmbCulture.Name = "cmbCulture"
        Me.cmbCulture.Size = New System.Drawing.Size(121, 21)
        Me.cmbCulture.TabIndex = 2
        '
        ' Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.cmbCulture)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents cmbCulture As ComboBox
End Class.



Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms

Public Class Form1
    Public Sub New()
        ' Initialize the form components
        InitializeComponent()
        ' Load the culture from environment variables or system and apply it
        InitializeCulture()
        ' Load localized resources manually
        LoadLocalizedResources()
        ' Set up event handler for culture change
        AddHandler cmbCulture.SelectedIndexChanged, AddressOf Me.CultureChanged
    End Sub

    Private Sub InitializeCulture()
        Dim culture As String = ResourceHelper.GetCultureFromEnvironment()
        ResourceHelper.SetCulture(culture)
        ' Pre-select the current culture in the ComboBox
        cmbCulture.SelectedItem = culture
    End Sub

    Private Sub LoadLocalizedResources()
        ' Use ResourceHelper to apply resources
        ResourceHelper.ApplyResources(Me, "Form1")
    End Sub

    Private Sub CultureChanged(sender As Object, e As EventArgs)
        Dim selectedCulture As String = cmbCulture.SelectedItem.ToString()
        If Not String.IsNullOrEmpty(selectedCulture) Then
            ResourceHelper.SaveCultureToEnvironment(selectedCulture)
            ResourceHelper.SetCulture(selectedCulture)
            ' Reload the form to apply new culture settings
            Me.Controls.Clear()
            InitializeComponent()
            InitializeCulture()
            LoadLocalizedResources()
            ' Set up event handler again after reinitializing components
            AddHandler cmbCulture.SelectedIndexChanged, AddressOf Me.CultureChanged
        End If
    End Sub
End Class
Imports System.Globalization
Imports System.IO
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Const CultureEnvironmentVariable As String = "MyApplicationCulture"

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        Dim culture As String = Thread.CurrentThread.CurrentUICulture.Name
        Dim resourceFileName As String = Path.Combine(Application.StartupPath, $"Resources\{resourceBaseName}.{culture}.resx")

        ' Fall back to default resource file if specific culture file doesn't exist
        If Not File.Exists(resourceFileName) Then
            resourceFileName = Path.Combine(Application.StartupPath, $"Resources\{resourceBaseName}.resx")
        End If

        ' Load resources from the specified .resx file
        Using resourceManager As New ResXResourceReader(resourceFileName)
            For Each entry As DictionaryEntry In resourceManager
                Dim resourceKey As String = entry.Key.ToString()
                Dim resourceValue As String = entry.Value.ToString()
                ApplyResourceToControl(form, resourceKey, resourceValue)
            Next
        End Using
    End Sub

    Private Shared Sub ApplyResourceToControl(form As Form, resourceKey As String, resourceValue As String)
        Dim controlProperty = resourceKey.Split("."c)
        If controlProperty.Length = 2 Then
            Dim controlName = controlProperty(0)
            Dim propertyName = controlProperty(1)

            Dim control = FindControlRecursive(form, controlName)
            If control IsNot Nothing Then
                Dim prop = control.GetType().GetProperty(propertyName)
                If prop IsNot Nothing AndAlso prop.CanWrite Then
                    prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                End If
            End If
        End If
    End Sub

    Private Shared Function FindControlRecursive(parent As Control, controlName As String) As Control
        If parent.Name = controlName Then
            Return parent
        End If

        For Each ctrl As Control In parent.Controls
            Dim foundCtrl = FindControlRecursive(ctrl, controlName)
            If foundCtrl IsNot Nothing Then
                Return foundCtrl
            End If
        Next

        Return Nothing
    End Function

    Public Shared Function GetCultureFromEnvironment() As String
        Dim culture As String = Environment.GetEnvironmentVariable(CultureEnvironmentVariable, EnvironmentVariableTarget.User)
        If String.IsNullOrEmpty(culture) Then
            culture = CultureInfo.CurrentUICulture.Name
            SaveCultureToEnvironment(culture)
        End If
        Return culture
    End Function

    Public Shared Sub SaveCultureToEnvironment(culture As String)
        Environment.SetEnvironmentVariable(CultureEnvironmentVariable, culture, EnvironmentVariableTarget.User)
    End Sub

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class


Using Resouree manager class

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Const CultureEnvironmentVariable As String = "MyApplicationCulture"
    Private Shared resourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        Dim culture As String = Thread.CurrentThread.CurrentUICulture.Name
        Dim assembly As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        resourceManager = New ResourceManager($"{assembly.GetName().Name}.Resources.{resourceBaseName}", assembly)

        For Each control As Control In GetAllControls(form)
            ApplyResourceToControl(control)
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control)
        Dim resourceName As String = control.Name.Replace("_", ".")
        Dim resourceString As String = resourceManager.GetString(resourceName)
        If Not String.IsNullOrEmpty(resourceString) Then
            If TypeOf control Is Label Then
                DirectCast(control, Label).Text = resourceString
            ElseIf TypeOf control Is Button Then
                DirectCast(control, Button).Text = resourceString
            ElseIf TypeOf control Is TextBox Then
                DirectCast(control, TextBox).Text = resourceString
            ElseIf TypeOf control Is ComboBox Then
                Dim comboBox As ComboBox = DirectCast(control, ComboBox)
                Dim index As Integer = comboBox.FindStringExact(resourceString)
                If index <> -1 Then
                    comboBox.SelectedIndex = index
                End If
            ElseIf TypeOf control Is ToolTip Then
                DirectCast(control, ToolTip).SetToolTip(control, resourceString)
            End If
        End If

        For Each childControl As Control In control.Controls
            ApplyResourceToControl(childControl)
        Next
    End Sub

    Private Shared Function GetAllControls(parent As Control) As IEnumerable(Of Control)
        Dim controls = parent.Controls.Cast(Of Control)()
        Return controls.SelectMany(Function(ctrl) GetAllControls(ctrl)).Concat(controls)
    End Function

    Public Shared Function GetCultureFromEnvironment() As String
        Dim culture As String = Environment.GetEnvironmentVariable(CultureEnvironmentVariable, EnvironmentVariableTarget.User)
        If String.IsNullOrEmpty(culture) Then
            culture = CultureInfo.CurrentUICulture.Name
            SaveCultureToEnvironment(culture)
        End If
        Return culture
    End Function

    Public Shared Sub SaveCultureToEnvironment(culture As String)
        Environment.SetEnvironmentVariable(CultureEnvironmentVariable, culture, EnvironmentVariableTarget.User)
    End Sub

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class


====just to modify

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        Dim culture As String = Thread.CurrentThread.CurrentUICulture.Name
        Dim resourceFileName As String = Path.Combine(Application.StartupPath, $"Resources\{resourceBaseName}.{culture}.resx")

        ' Fall back to default resource file if specific culture file doesn't exist
        If Not File.Exists(resourceFileName) Then
            resourceFileName = Path.Combine(Application.StartupPath, $"Resources\{resourceBaseName}.resx")
        End If

        ' Load resources from the specified .resx file
        Using resourceManager As New ResXResourceReader(resourceFileName)
            For Each entry As DictionaryEntry In resourceManager
                Dim resourceKey As String = entry.Key.ToString()
                Dim resourceValue As String = entry.Value.ToString()
                ApplyResourceToControl(form, resourceKey, resourceValue)
            Next
        End Using
    End Sub

    Private Shared Sub ApplyResourceToControl(form As Form, resourceKey As String, resourceValue As String)
        Dim controlProperty = resourceKey.Split("."c)
        If controlProperty.Length = 2 Then
            Dim controlName = controlProperty(0)
            Dim propertyName = controlProperty(1)

            Dim control = FindControlRecursive(form, controlName)
            If control IsNot Nothing Then
                Dim prop = control.GetType().GetProperty(propertyName)
                If prop IsNot Nothing AndAlso prop.CanWrite Then
                    prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                End If
            End If
        End If
    End Sub

    Private Shared Function FindControlRecursive(parent As Control, controlName As String) As Control
        If parent.Name = controlName Then
            Return parent
        End If

        For Each ctrl As Control In parent.Controls
            Dim foundCtrl = FindControlRecursive(ctrl, controlName)
            If foundCtrl IsNot Nothing Then
                Return foundCtrl
            End If
        Next

        Return Nothing
    End Function

    ===using resource manager 

    Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.IO

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl)
            ' Recursively apply resources to child controls
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            ' Retrieve the localized resource string from the resource file
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Replace placeholders in the resource string with the provided values
                Return String.Format(localizedResourceString, values)
            Else
                ' Handle missing resource string
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class


====using 

Imports System.Globalization
Imports System.Resources
Imports System.Threading

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ' Include the Resources folder in the base name
        ResourceManager = New ResourceManager("Resources." & resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parent.Name)
            ' Recursively apply resources to child controls
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            ' Retrieve the localized resource string from the resource file
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Replace placeholders in the resource string with the provided values
                Return String.Format(localizedResourceString, values)
            Else
                ' Handle missing resource string
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
==============

Imports System.Globalization
Imports System.Resources

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Public Shared Function GetAllStrings(culture As CultureInfo) As Dictionary(Of String, String)
        Dim allStrings As New Dictionary(Of String, String)()

        Dim resourceSet As ResourceSet = ResourceManager.GetResourceSet(culture, createIfNotExists:=True, tryParents:=True)
        If resourceSet IsNot Nothing Then
            For Each resourceKey As Object In resourceSet.Keys
                Dim resourceValue As Object = resourceSet.GetObject(resourceKey)
                allStrings.Add(resourceKey.ToString(), If(resourceValue IsNot Nothing, resourceValue.ToString(), ""))
            Next
        End If

        Return allStrings
    End Function
End Class

=================


Imports System.Globalization
Imports System.Resources
Imports System.Threading

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parent.Name)
            ' Recursively apply resources to child controls
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            ' Retrieve the localized resource string from the resource file
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Replace placeholders in the resource string with the provided values
                Return String.Format(localizedResourceString, values)
            Else
                ' Handle missing resource string
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class


-----

Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ResourceHelper.SetCulture("fr-CA") ' Set to desired culture
    ResourceHelper.ApplyResources(Me, "Resources") ' Apply resources to the form using the single resource file
End Sub

=======================

using menu strip

Imports System.Globalization
Imports System.Resources
Imports System.Threading

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ' Apply resources to the form itself
        ApplyResourceToForm(form)
        ' Apply resources to all controls on the form
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        Dim formProperties = form.GetType().GetProperties()
        For Each prop In formProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{form.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(form, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on form {form.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            ' Recursively apply resources to child controls
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next
        ' Apply resources to MenuStrip items if present
        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        ' Recursively apply resources to dropdown items
        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            ' Retrieve the localized resource string from the resource file
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Replace placeholders in the resource string with the provided values
                Return String.Format(localizedResourceString, values)
            Else
                ' Handle missing resource string
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class

====
Imports System.Globalization
Imports System.Resources
Imports System.Threading

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ' Apply resources to the form itself
        ApplyResourceToForm(form)
        ' Apply resources to all controls on the form
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        Dim formProperties = form.GetType().GetProperties()
        For Each prop In formProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{form.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(form, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on form {form.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            ' Recursively apply resources to child controls
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        ' Apply resources to special container controls like MenuStrip
        If TypeOf parent Is MenuStrip Then
            ApplyResourcesToMenuStrip(CType(parent, MenuStrip), parentFormName)
        ElseIf TypeOf parent Is ToolStrip Then
            ApplyResourcesToToolStrip(CType(parent, ToolStrip), parentFormName)
        ElseIf TypeOf parent Is ToolStripDropDownButton Then
            ApplyResourcesToToolStripDropDownButton(CType(parent, ToolStripDropDownButton), parentFormName)
        ElseIf TypeOf parent Is ContextMenuStrip Then
            ApplyResourcesToContextMenuStrip(CType(parent, ContextMenuStrip), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToToolStrip(toolStrip As ToolStrip, parentFormName As String)
        For Each item As ToolStripItem In toolStrip.Items
            If TypeOf item Is ToolStripDropDownItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripDropDownItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToToolStripDropDownButton(toolStripDropDownButton As ToolStripDropDownButton, parentFormName As String)
        For Each item As ToolStripItem In toolStripDropDownButton.DropDownItems
            If TypeOf item Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToContextMenuStrip(contextMenuStrip As ContextMenuStrip, parentFormName As String)
        For Each item As ToolStripItem In contextMenuStrip.Items
            If TypeOf item Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        ' Recursively apply resources to dropdown items
        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next

        ' Handle special controls
        If TypeOf control Is GroupBox Then
            ApplyResourcesToControls(CType(control, GroupBox), parentFormName)
        ElseIf TypeOf control Is TabControl Then
            ApplyResourcesToTabControl(CType(control, TabControl), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourcesToTabControl(tabControl As TabControl, parentFormName As String)
        For Each tabPage As TabPage In tabControl.TabPages
            ApplyResourceToControl(tabPage, parentFormName)
            ApplyResourcesToControls(tabPage, parentFormName)
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            ' Retrieve the localized resource string from the resource file
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Replace placeholders in the resource string with the provided values
                Return String.Format(localizedResourceString, values)
            Else
                ' Handle missing resource string
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
====

Imports System.Globalization
Imports System.Reflection
Imports System.Resources
Imports System.Threading

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared PropertyCache As New Dictionary(Of Type, PropertyInfo())

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToControl(form, form.Name)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToControl(control As Control, parentName As String)
        Dim controlType = control.GetType()

        ' Get properties of the control type, cached for performance
        Dim properties As PropertyInfo() = GetProperties(controlType)

        ' Apply resources to the control's string properties
        For Each prop In properties
            If prop.CanWrite AndAlso prop.PropertyType Is GetType(String) Then
                Dim resourceKey As String = $"{parentName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, resourceValue, Nothing)
                    End If
                Catch ex As Exception
                    ' Log or handle the error
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next

        ' Recursively apply resources to child controls
        For Each child As Control In control.Controls
            ApplyResourcesToControl(child, parentName)
        Next

        ' Handle special control types
        If TypeOf control Is MenuStrip Then
            ApplyResourcesToMenuStrip(CType(control, MenuStrip), parentName)
        ElseIf TypeOf control Is ToolStrip Then
            ApplyResourcesToToolStrip(CType(control, ToolStrip), parentName)
        ElseIf TypeOf control Is ContextMenuStrip Then
            ApplyResourcesToContextMenuStrip(CType(control, ContextMenuStrip), parentName)
        ElseIf TypeOf control Is TabControl Then
            ApplyResourcesToTabControl(CType(control, TabControl), parentName)
        ElseIf TypeOf control Is GroupBox Then
            ApplyResourcesToGroupBox(CType(control, GroupBox), parentName)
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToToolStripItem(menuItem, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToToolStrip(toolStrip As ToolStrip, parentName As String)
        For Each item As ToolStripItem In toolStrip.Items
            ApplyResourceToToolStripItem(item, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToContextMenuStrip(contextMenuStrip As ContextMenuStrip, parentName As String)
        For Each item As ToolStripItem In contextMenuStrip.Items
            ApplyResourceToToolStripItem(item, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToTabControl(tabControl As TabControl, parentName As String)
        For Each tabPage As TabPage In tabControl.TabPages
            ApplyResourcesToControl(tabPage, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToGroupBox(groupBox As GroupBox, parentName As String)
        ApplyResourcesToControl(groupBox, parentName)
    End Sub

    Private Shared Sub ApplyResourceToToolStripItem(item As ToolStripItem, parentName As String)
        Dim resourceKey As String = $"{parentName}.{item.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                item.Text = resourceValue
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error setting text for tool strip item {item.Name}: {ex.Message}")
        End Try

        ' Recursively apply resources to dropdown items if applicable
        If TypeOf item Is ToolStripDropDownItem Then
            For Each dropDownItem As ToolStripItem In CType(item, ToolStripDropDownItem).DropDownItems
                ApplyResourceToToolStripItem(dropDownItem, parentName)
            Next
        End If
    End Sub

    Private Shared Function GetProperties(type As Type) As PropertyInfo()
        If Not PropertyCache.ContainsKey(type) Then
            PropertyCache(type) = type.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        End If
        Return PropertyCache(type)
    End Function

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            ' Retrieve the localized resource string from the resource file
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Replace placeholders in the resource string with the provided values
                Return String.Format(localizedResourceString, values)
            Else
                ' Handle missing resource string
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log or handle the error
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
==============
with form heading

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Diagnostics
Imports System.Collections.Concurrent
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared PropertyCache As New ConcurrentDictionary(Of Type, PropertyInfo())

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToComponent(form, form.Name)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToComponent(component As Component, parentName As String)
        If TypeOf component Is Form OrElse TypeOf component Is Control Then
            ApplyResourceToProperties(component, parentName)
            If TypeOf component Is Control Then
                For Each child As Control In CType(component, Control).Controls
                    ApplyResourcesToComponent(child, parentName)
                Next
            End If
        End If

        Select Case True
            Case TypeOf component Is MenuStrip
                ApplyResourcesToMenuStrip(CType(component, MenuStrip), parentName)
            Case TypeOf component Is ToolStrip
                ApplyResourcesToToolStrip(CType(component, ToolStrip), parentName)
            Case TypeOf component Is ContextMenuStrip
                ApplyResourcesToContextMenuStrip(CType(component, ContextMenuStrip), parentName)
            Case TypeOf component Is StatusStrip
                ApplyResourcesToStatusStrip(CType(component, StatusStrip), parentName)
            Case TypeOf component Is SplitContainer
                ApplyResourcesToSplitContainer(CType(component, SplitContainer), parentName)
            Case TypeOf component Is TreeView
                ApplyResourcesToTreeView(CType(component, TreeView), parentName)
            Case TypeOf component Is ListView
                ApplyResourcesToListView(CType(component, ListView), parentName)
            Case TypeOf component Is ComboBox
                ApplyResourcesToComboBox(CType(component, ComboBox), parentName)
        End Select
    End Sub

    Private Shared Sub ApplyResourceToProperties(component As Component, parentName As String)
        Dim componentType = component.GetType()
        Dim componentProperties = PropertyCache.GetOrAdd(componentType, Function(t) t.GetProperties())

        For Each prop In componentProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentName}.{component.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(component, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on component {component.Name}: {ex.Message}")
                End Try
            End If
        Next

        ' Special handling for form's Text property (form heading)
        If TypeOf component Is Form Then
            Dim form As Form = CType(component, Form)
            Dim formResourceKey As String = $"{parentName}.{form.Name}.Text"
            Try
                Dim formResourceValue As String = ResourceManager.GetString(formResourceKey, Thread.CurrentThread.CurrentUICulture)
                If formResourceValue IsNot Nothing Then
                    form.Text = formResourceValue
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting form text for form {form.Name}: {ex.Message}")
            End Try
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToToolStrip(toolStrip As ToolStrip, parentName As String)
        For Each item As ToolStripItem In toolStrip.Items
            If TypeOf item Is ToolStripDropDownItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripDropDownItem), parentName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToContextMenuStrip(contextMenuStrip As ContextMenuStrip, parentName As String)
        For Each item As ToolStripItem In contextMenuStrip.Items
            If TypeOf item Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripMenuItem), parentName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToStatusStrip(statusStrip As StatusStrip, parentName As String)
        For Each item As ToolStripItem In statusStrip.Items
            If TypeOf item Is ToolStripStatusLabel Then
                ApplyResourceToProperties(item, parentName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToSplitContainer(splitContainer As SplitContainer, parentName As String)
        ApplyResourcesToComponent(splitContainer.Panel1, parentName)
        ApplyResourcesToComponent(splitContainer.Panel2, parentName)
    End Sub

    Private Shared Sub ApplyResourcesToTreeView(treeView As TreeView, parentName As String)
        For Each node As TreeNode In treeView.Nodes
            ApplyResourceToTreeNode(node, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToTreeNode(node As TreeNode, parentName As String)
        Dim resourceKey As String = $"{parentName}.{node.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                node.Text = resourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for tree node {node.Name}: {ex.Message}")
        End Try

        For Each childNode As TreeNode In node.Nodes
            ApplyResourceToTreeNode(childNode, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToListView(listView As ListView, parentName As String)
        For Each item As ListViewItem In listView.Items
            ApplyResourceToProperties(item, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToComboBox(comboBox As ComboBox, parentName As String)
        For i As Integer = 0 To comboBox.Items.Count - 1
            Dim item = combo
======without form heading

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Diagnostics
Imports System.Collections.Concurrent

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared PropertyCache As New ConcurrentDictionary(Of Type, PropertyInfo())

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToComponent(form, form.Name)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToComponent(component As Component, parentName As String)
        If TypeOf component Is Form OrElse TypeOf component Is Control Then
            ApplyResourceToProperties(component, parentName)
            If TypeOf component Is Control Then
                For Each child As Control In CType(component, Control).Controls
                    ApplyResourcesToComponent(child, parentName)
                Next
            End If
        End If

        Select Case True
            Case TypeOf component Is MenuStrip
                ApplyResourcesToMenuStrip(CType(component, MenuStrip), parentName)
            Case TypeOf component Is ToolStrip
                ApplyResourcesToToolStrip(CType(component, ToolStrip), parentName)
            Case TypeOf component Is ContextMenuStrip
                ApplyResourcesToContextMenuStrip(CType(component, ContextMenuStrip), parentName)
            Case TypeOf component Is StatusStrip
                ApplyResourcesToStatusStrip(CType(component, StatusStrip), parentName)
            Case TypeOf component Is SplitContainer
                ApplyResourcesToSplitContainer(CType(component, SplitContainer), parentName)
            Case TypeOf component Is TreeView
                ApplyResourcesToTreeView(CType(component, TreeView), parentName)
            Case TypeOf component Is ListView
                ApplyResourcesToListView(CType(component, ListView), parentName)
            Case TypeOf component Is ComboBox
                ApplyResourcesToComboBox(CType(component, ComboBox), parentName)
        End Select
    End Sub

    Private Shared Sub ApplyResourceToProperties(component As Component, parentName As String)
        Dim componentType = component.GetType()
        Dim componentProperties = PropertyCache.GetOrAdd(componentType, Function(t) t.GetProperties())

        For Each prop In componentProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentName}.{component.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(component, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on component {component.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToToolStrip(toolStrip As ToolStrip, parentName As String)
        For Each item As ToolStripItem In toolStrip.Items
            If TypeOf item Is ToolStripDropDownItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripDropDownItem), parentName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToContextMenuStrip(contextMenuStrip As ContextMenuStrip, parentName As String)
        For Each item As ToolStripItem In contextMenuStrip.Items
            If TypeOf item Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(item, ToolStripMenuItem), parentName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToStatusStrip(statusStrip As StatusStrip, parentName As String)
        For Each item As ToolStripItem In statusStrip.Items
            If TypeOf item Is ToolStripStatusLabel Then
                ApplyResourceToProperties(item, parentName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToSplitContainer(splitContainer As SplitContainer, parentName As String)
        ApplyResourcesToComponent(splitContainer.Panel1, parentName)
        ApplyResourcesToComponent(splitContainer.Panel2, parentName)
    End Sub

    Private Shared Sub ApplyResourcesToTreeView(treeView As TreeView, parentName As String)
        For Each node As TreeNode In treeView.Nodes
            ApplyResourceToTreeNode(node, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToTreeNode(node As TreeNode, parentName As String)
        Dim resourceKey As String = $"{parentName}.{node.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                node.Text = resourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for tree node {node.Name}: {ex.Message}")
        End Try

        For Each childNode As TreeNode In node.Nodes
            ApplyResourceToTreeNode(childNode, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToListView(listView As ListView, parentName As String)
        For Each item As ListViewItem In listView.Items
            ApplyResourceToProperties(item, parentName)
        Next
    End Sub

    Private Shared Sub ApplyResourcesToComboBox(comboBox As ComboBox, parentName As String)
        For i As Integer = 0 To comboBox.Items.Count - 1
            Dim item = comboBox.Items(i)
            If TypeOf item Is String Then
                Dim resourceKey As String = $"{parentName}.{comboBox.Name}.Items[{i}]"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        comboBox.Items(i) = resourceValue
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting text for combobox item at index {i}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentName As String)
        Dim resourceKey As String = $"{parentName}.{menuItem.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentName)
            End If
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                Return String.Format(localizedResourceString, values)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
=================================Additional changes

Imports System.Collections.Generic
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared _resourceDictionary As Dictionary(Of String, Dictionary(Of String, String)) = New Dictionary(Of String, Dictionary(Of String, String))()

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        Try
            InitializeResourceManager(resourceBaseName)
            ApplyResourcesToControl(form)
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error applying resources: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ' Initialize resource manager with resource files based on base name
        ' For example, load "resourceBaseName.en-US.resx" for English (United States)
        ' and "resourceBaseName.fr-FR.resx" for French (France)
    End Sub

    Private Shared Sub ApplyResourcesToControl(control As Control)
        Try
            Dim resourceName As String = control.Name

            ' Check if resource exists for the control
            If _resourceDictionary.ContainsKey(resourceName) Then
                Dim resourceMappings = _resourceDictionary(resourceName)
                For Each kvp As KeyValuePair(Of String, String) In resourceMappings
                    Dim propertyName As String = kvp.Key
                    Dim propertyValue As String = kvp.Value
                    SetProperty(control, propertyName, propertyValue)
                Next
            End If

            ' Recursively apply resources to child controls
            If TypeOf control Is MenuStrip Then
                ApplyResourcesToMenuStrip(DirectCast(control, MenuStrip))
            ElseIf TypeOf control Is GroupBox Then
                Dim groupBox As GroupBox = DirectCast(control, GroupBox)
                For Each childControl As Control In groupBox.Controls
                    ApplyResourcesToControl(childControl)
                Next
            Else
                For Each childControl As Control In control.Controls
                    ApplyResourcesToControl(childControl)
                Next
            End If
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error applying resources to control {control.Name}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Sub SetProperty(control As Control, propertyName As String, propertyValue As String)
        Try
            If propertyName = "Text" Then
                control.Text = propertyValue
            ElseIf propertyName = "Visible" Then
                control.Visible = Boolean.Parse(propertyValue)
            ' Add more property handlers as needed
            End If
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error setting property {propertyName} for control {control.Name}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip)
        Try
            For Each menuItem As ToolStripMenuItem In menuStrip.Items
                ApplyResourceToMenuStripItem(menuItem)
            Next
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error applying resources to MenuStrip: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Sub ApplyResourceToMenuStripItem(menuItem As ToolStripMenuItem)
        Try
            Dim resourceName As String = menuItem.Name

            ' Check if resource exists for the menu item
            If _resourceDictionary.ContainsKey(resourceName) Then
                Dim resourceMappings = _resourceDictionary(resourceName)
                If resourceMappings.TryGetValue("Text", Dim propertyValue As String) Then
                    menuItem.Text = propertyValue
                End If
            End If

            ' Recursively apply resources to dropdown items
            For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
                If TypeOf dropDownItem Is ToolStripMenuItem Then
                    ApplyResourceToMenuStripItem(DirectCast(dropDownItem, ToolStripMenuItem))
                End If
            Next
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error applying resources to MenuStrip item {menuItem.Name}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Shared Sub AddResourceMapping(controlName As String, propertyName As String, propertyValue As String)
        Try
            If Not _resourceDictionary.ContainsKey(controlName) Then
                _resourceDictionary.Add(controlName, New Dictionary(Of String, String)())
            End If

            Dim resourceMappings = _resourceDictionary(controlName)
            If Not resourceMappings.ContainsKey(propertyName) Then
                resourceMappings.Add(propertyName, propertyValue)
            Else
                resourceMappings(propertyName) = propertyValue
            End If
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error adding resource mapping for control {controlName}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class

======================================
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Resources
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared _resourceManager As ResourceManager
    Private Shared _resourceDictionary As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        Try
            InitializeResourceManager(resourceBaseName)
            ApplyResourcesToControl(form)
        Catch ex As Exception
            ' Handle the exception, e.g., logging or displaying an error message
            MessageBox.Show($"Error applying resources: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ' Initialize resource manager with resource base name
        _resourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)

        ' Load all resources into a dictionary for faster access
        _resourceDictionary = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))()
        Dim cultures As CultureInfo() = CultureInfo.GetCultures(CultureTypes.AllCultures)
        For Each culture As CultureInfo In cultures
            Dim resourceSet As New Dictionary(Of String, Dictionary(Of String, String))()
            Dim resourceSetExists As Boolean = False
            For Each resourceKey As String In _resourceManager.GetResourceSet(culture, True, True).OfType(Of String)()
                resourceSet(resourceKey) = LoadResourceSetForControl(resourceKey, culture)
                resourceSetExists = True
            Next
            If resourceSetExists Then
                _resourceDictionary(culture.Name) = resourceSet
            End If
        Next
    End Sub

    Private Shared Function LoadResourceSetForControl(resourceKey As String, culture As CultureInfo) As Dictionary(Of String, String)
        Dim resourceSet As New Dictionary(Of String, String)()
        Dim parts As String() = resourceKey.Split("."c)

        ' Ensure that the resource key follows the pattern {parentFormName}.{control.Name}.{prop.Name}
        If parts.Length = 3 Then
            Dim parentFormName As String = parts(0)
            Dim controlName As String = parts(1)
            Dim propertyName As String = parts(2)
            Dim resourceValue As String = _resourceManager.GetString(resourceKey, culture)

            If Not String.IsNullOrEmpty(resourceValue) Then
                resourceSet(propertyName) = resourceValue
            End If
        End If

        Return resourceSet
    End Function

    Private Shared Sub ApplyResourcesToControl(control As Control)
        Dim cultureName As String = CultureInfo.CurrentUICulture.Name

        ' Check if resource exists for the current culture
        If _resourceDictionary.ContainsKey(cultureName) Then
            Dim resourceSet As Dictionary(Of String, String) = GetResourceSetForControl(control.Name, cultureName)

            ' Apply resources to the control
            ApplyResourceProperties(control, resourceSet)
        End If

        ' Recursively apply resources to child controls
        For Each childControl As Control In control.Controls
            ApplyResourcesToControl(childControl)
        Next
    End Sub

    Private Shared Function GetResourceSetForControl(controlName As String, cultureName As String) As Dictionary(Of String, String)
        Dim resourceSet As New Dictionary(Of String, String)()

        If _resourceDictionary(cultureName).ContainsKey(controlName) Then
            resourceSet = _resourceDictionary(cultureName)(controlName)
        End If

        Return resourceSet
    End Function

    Private Shared Sub ApplyResourceProperties(control As Control, resourceSet As Dictionary(Of String, String))
        For Each kvp As KeyValuePair(Of String, String) In resourceSet
            Dim propertyName As String = kvp.Key
            Dim propertyValue As String = kvp.Value
            SetProperty(control, propertyName, propertyValue)
        Next
    End Sub

    Private Shared Sub SetProperty(control As Control, propertyName As String, propertyValue As String)
        Select Case propertyName
            Case "Text"
                control.Text = propertyValue
            Case "Visible"
                control.Visible = Boolean.Parse(propertyValue)
            ' Add more property handlers as needed
            ' For example:
            ' Case "Enabled"
            '     control.Enabled = Boolean.Parse(propertyValue)
        End Select
    End Sub
End Class

==using combobox

Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
    InitializeResourceManager(resourceBaseName)
    ApplyResourcesToFormAndControls(form)
End Sub

Private Shared Sub InitializeResourceManager(resourceBaseName As String)
    ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
End Sub

Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
    ApplyResourceToForm(form)
    ApplyResourcesToControls(form, form.Name)
End Sub

Private Shared Sub ApplyResourceToForm(form As Form)
    Dim formProperties = form.GetType().GetProperties()
    For Each prop In formProperties
        If prop.CanWrite Then
            Dim resourceKey As String = $"{form.Name}.{prop.Name}"
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                If resourceValue IsNot Nothing Then
                    prop.SetValue(form, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting property {prop.Name} on form {form.Name}: {ex.Message}")
            End Try
        End If
    Next
End Sub

Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
    For Each ctrl As Control In parent.Controls
        ApplyResourceToControl(ctrl, parentFormName)
        If ctrl.HasChildren Then
            ApplyResourcesToControls(ctrl, parentFormName)
        End If
    Next
    If TypeOf parent Is Form Then
        Dim form As Form = CType(parent, Form)
        For Each ctrl As Control In form.Controls
            If TypeOf ctrl Is MenuStrip Then
                ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
            End If
        Next
    End If
End Sub

Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
    For Each menuItem As ToolStripMenuItem In menuStrip.Items
        ApplyResourceToMenuItem(menuItem, parentFormName)
    Next
End Sub

Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
    Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
    Try
        Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
        If resourceValue IsNot Nothing Then
            menuItem.Text = resourceValue
        End If
    Catch ex As Exception
        Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
    End Try

    For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
        If TypeOf dropDownItem Is ToolStripMenuItem Then
            ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
        End If
    Next
End Sub

Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
    Dim controlProperties = control.GetType().GetProperties()
    For Each prop In controlProperties
        If prop.CanWrite Then
            Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                If resourceValue IsNot Nothing Then
                    prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
            End Try
        End If
    Next

    ' Handle ComboBox items
    If TypeOf control Is ComboBox Then
        ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
    End If
End Sub

Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
    Dim items = comboBox.Items
    comboBox.Items.Clear()
    For i As Integer = 0 To items.Count - 1
        Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                comboBox.Items.Add(resourceValue)
            Else
                comboBox.Items.Add(items(i)) ' Add original item if no resource is found
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
            comboBox.Items.Add(items(i)) ' Add original item if an error occurs
        End Try
    Next
End Sub

Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
    Try
        Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
        If localizedResourceString IsNot Nothing Then
            Return String.Format(localizedResourceString, values)
        Else
            Return $"[Missing Resource: {resourceKey}]"
        End If
    Catch ex As Exception
        Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
        Return $"[Error: {resourceKey}]"
    End Try
End Function

Public Shared Sub SetCulture(cultureName As String)
    Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
End Sub








============================combobox and menustrip

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ApplyResourceToForm(form)
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        Dim formProperties = form.GetType().GetProperties()
        For Each prop In formProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{form.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(form, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on form {form.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next

        ' Handle ComboBox items
        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i)) ' Add original item if no resource is found
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i)) ' Add original item if an error occurs
            End Try
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                Return String.Format(localizedResourceString, values)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class

=====================================

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ApplyResourceToForm(form)
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        Dim formProperties = form.GetType().GetProperties()
        For Each prop In formProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{form.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(form, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on form {form.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next

        ' Handle ComboBox items and Text property
        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
        ' Localize the Text property of the ComboBox
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Try
            Dim textResourceValue As String = ResourceManager.GetString(textResourceKey, Thread.CurrentThread.CurrentUICulture)
            If textResourceValue IsNot Nothing Then
                comboBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for ComboBox {comboBox.Name}: {ex.Message}")
        End Try

        ' Localize the items of the ComboBox
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i)) ' Add original item if no resource is found
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i)) ' Add original item if an error occurs
            End Try
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                Return String.Format(localizedResourceString, values)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
============================== performative code
Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared ResourceCache As New Dictionary(Of String, String)

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ApplyResourceToForm(form)
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        ApplyResourceToObject(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        ApplyResourceToObject(menuItem, $"{parentFormName}.{menuItem.Name}")
        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        ApplyResourceToObject(control, $"{parentFormName}.{control.Name}")

        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourceToObject(obj As Object, baseResourceKey As String)
        Dim objProperties = obj.GetType().GetProperties()
        For Each prop In objProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{baseResourceKey}.{prop.Name}"
                Try
                    Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(obj, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on {obj.GetType().Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Try
            Dim textResourceValue As String = GetLocalizedResourceByKey(textResourceKey)
            If textResourceValue IsNot Nothing Then
                comboBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for ComboBox {comboBox.Name}: {ex.Message}")
        End Try

        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i))
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i))
            End Try
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        If ResourceCache.ContainsKey(resourceKey) Then
            Return ResourceCache(resourceKey)
        End If

        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ResourceCache(resourceKey) = localizedResourceString
                Return If(values.Length > 0, String.Format(localizedResourceString, values), localizedResourceString)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
        ResourceCache.Clear()
    End Sub
End Class

=============================== more perfomatic

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared ResourceCache As New Dictionary(Of String, String)

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ApplyResourceToForm(form)
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        ApplyResourceToObject(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        ApplyResourceToObject(menuItem, $"{parentFormName}.{menuItem.Name}")
        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        ApplyResourceToObject(control, $"{parentFormName}.{control.Name}")

        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourceToObject(obj As Object, baseResourceKey As String)
        Dim objProperties = obj.GetType().GetProperties()
        For Each prop In objProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{baseResourceKey}.{prop.Name}"
                Try
                    Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(obj, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on {obj.GetType().Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
        ' Localize the Text property of the ComboBox
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Try
            Dim textResourceValue As String = GetLocalizedResourceByKey(textResourceKey)
            If textResourceValue IsNot Nothing Then
                comboBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for ComboBox {comboBox.Name}: {ex.Message}")
        End Try

        ' Localize the items of the ComboBox
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i)) ' Add original item if no resource is found
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i)) ' Add original item if an error occurs
            End Try
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        If ResourceCache.ContainsKey(resourceKey) Then
            Return If(values.Length > 0, String.Format(ResourceCache(resourceKey), values), ResourceCache(resourceKey))
        End If

        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ResourceCache(resourceKey) = localizedResourceString
                Return If(values.Length > 0, String.Format(localizedResourceString, values), localizedResourceString)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
        ResourceCache.Clear()
    End Sub
End Class

===same but with detila explanation

Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared ResourceCache As New Dictionary(Of String, String)

    ''' <summary>
    ''' Apply localization resources to the specified form and its controls.
    ''' </summary>
    ''' <param name="form">The form to localize.</param>
    ''' <param name="resourceBaseName">The base name of the resource file.</param>
    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    ''' <summary>
    ''' Initialize the resource manager with the specified base name.
    ''' </summary>
    ''' <param name="resourceBaseName">The base name of the resource file.</param>
    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    ''' <summary>
    ''' Apply resources to the form and its controls.
    ''' </summary>
    ''' <param name="form">The form to localize.</param>
    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ApplyResourceToForm(form)
        ApplyResourcesToControls(form, form.Name)
    End Sub

    ''' <summary>
    ''' Apply resources to the form itself.
    ''' </summary>
    ''' <param name="form">The form to localize.</param>
    Private Shared Sub ApplyResourceToForm(form As Form)
        ApplyResourceToObject(form, form.Name)
    End Sub

    ''' <summary>
    ''' Apply resources to all controls within the specified parent control.
    ''' </summary>
    ''' <param name="parent">The parent control containing child controls to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Apply resources to a MenuStrip and its items.
    ''' </summary>
    ''' <param name="menuStrip">The MenuStrip to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    ''' <summary>
    ''' Apply resources to a ToolStripMenuItem and its sub-items.
    ''' </summary>
    ''' <param name="menuItem">The ToolStripMenuItem to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        ApplyResourceToObject(menuItem, $"{parentFormName}.{menuItem.Name}")
        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Apply resources to a control.
    ''' </summary>
    ''' <param name="control">The control to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        ApplyResourceToObject(control, $"{parentFormName}.{control.Name}")

        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
        End If
    End Sub

    ''' <summary>
    ''' Apply resources to a generic object (form, control, or menu item).
    ''' </summary>
    ''' <param name="obj">The object to localize.</param>
    ''' <param name="baseResourceKey">The base key for the resources.</param>
    Private Shared Sub ApplyResourceToObject(obj As Object, baseResourceKey As String)
        Dim objProperties = obj.GetType().GetProperties()
        For Each prop In objProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{baseResourceKey}.{prop.Name}"
                Try
                    Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(obj, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on {obj.GetType().Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    ''' <summary>
    ''' Apply resources to the items of a ComboBox.
    ''' </summary>
    ''' <param name="comboBox">The ComboBox to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
        ' Localize the Text property of the ComboBox
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Try
            Dim textResourceValue As String = GetLocalizedResourceByKey(textResourceKey)
            If textResourceValue IsNot Nothing Then
                comboBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for ComboBox {comboBox.Name}: {ex.Message}")
        End Try

        ' Localize the items of the ComboBox
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i)) ' Add original item if no resource is found
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i)) ' Add original item if an error occurs
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Retrieve a localized resource string by its key.
    ''' </summary>
    ''' <param name="resourceKey">The resource key.</param>
    ''' <param name="values">Optional values to format the resource string.</param>
    ''' <returns>The localized resource string.</returns>
    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        If ResourceCache.ContainsKey(resourceKey) Then
            Return If(values.Length > 0, String.Format(ResourceCache(resourceKey), values), ResourceCache(resourceKey))
        End If

        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ResourceCache(resourceKey) = localizedResourceString
                Return If(values.Length > 0, String.Format(localizedResourceString, values), localizedResourceString)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    ''' <summary>
    ''' Set the current UI culture.
    ''' </summary>
    ''' <param name="cultureName">The name of the culture to set.</param>
    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
        ResourceCache.Clear()
    End Sub
End Class

=====groupbox, combobox and menustrip support
Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager

    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        InitializeResourceManager(resourceBaseName)
        ApplyResourcesToFormAndControls(form)
    End Sub

    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ApplyResourceToForm(form)
        ApplyResourcesToControls(form, form.Name)
    End Sub

    Private Shared Sub ApplyResourceToForm(form As Form)
        Dim formProperties = form.GetType().GetProperties()
        For Each prop In formProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{form.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(form, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on form {form.Name}: {ex.Message}")
                End Try
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        For Each ctrl As Control In parent.Controls
            ApplyResourceToControl(ctrl, parentFormName)
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
        Try
            Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        Dim controlProperties = control.GetType().GetProperties()
        For Each prop In controlProperties
            If prop.CanWrite Then
                Dim resourceKey As String = $"{parentFormName}.{control.Name}.{prop.Name}"
                Try
                    Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                    If resourceValue IsNot Nothing Then
                        prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error setting property {prop.Name} on control {control.Name}: {ex.Message}")
                End Try
            End If
        Next

        ' Handle ComboBox items and Text property
        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBox(CType(control, ComboBox), parentFormName)
        ElseIf TypeOf control Is GroupBox Then
            ApplyResourcesToGroupBox(CType(control, GroupBox), parentFormName)
        End If
    End Sub

    Private Shared Sub ApplyResourcesToComboBox(comboBox As ComboBox, parentFormName As String)
        ' Localize the Text property of the ComboBox
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Try
            Dim textResourceValue As String = ResourceManager.GetString(textResourceKey, Thread.CurrentThread.CurrentUICulture)
            If textResourceValue IsNot Nothing Then
                comboBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for ComboBox {comboBox.Name}: {ex.Message}")
        End Try

        ' Localize the items of the ComboBox
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i)) ' Add original item if no resource is found
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i)) ' Add original item if an error occurs
            End Try
        Next
    End Sub

    Private Shared Sub ApplyResourcesToGroupBox(groupBox As GroupBox, parentFormName As String)
        ' Localize the Text property of the GroupBox
        Dim textResourceKey As String = $"{parentFormName}.{groupBox.Name}.Text"
        Try
            Dim textResourceValue As String = ResourceManager.GetString(textResourceKey, Thread.CurrentThread.CurrentUICulture)
            If textResourceValue IsNot Nothing Then
                groupBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for GroupBox {groupBox.Name}: {ex.Message}")
        End Try

        ' Apply resources to controls within the GroupBox recursively
        For Each ctrl As Control In groupBox.Controls
            ApplyResourceToControl(ctrl, $"{parentFormName}.{groupBox.Name}")
        Next
    End Sub

    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                Return String.Format(localizedResourceString, values)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
===perfomative code using resource cache
Imports System.Globalization
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    ' ResourceManager to access localized resources
    Private Shared ResourceManager As ResourceManager

    ' Dictionary to cache resource values for better performance
    Private Shared ResourceCache As New Dictionary(Of String, String)

    ''' <summary>
    ''' Apply localization resources to the specified form and its controls.
    ''' </summary>
    ''' <param name="form">The form to localize.</param>
    ''' <param name="resourceBaseName">The base name of the resource file.</param>
    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        ' Initialize the resource manager
        InitializeResourceManager(resourceBaseName)
        
        ' Apply resources to the form and its controls
        ApplyResourcesToFormAndControls(form)
    End Sub

    ''' <summary>
    ''' Initialize the resource manager with the specified base name.
    ''' </summary>
    ''' <param name="resourceBaseName">The base name of the resource file.</param>
    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ' Create a new ResourceManager instance with the specified base name
        ' and the assembly where the resource file is located
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
    End Sub

    ''' <summary>
    ''' Apply resources to the form and its controls.
    ''' </summary>
    ''' <param name="form">The form to localize.</param>
    Private Shared Sub ApplyResourcesToFormAndControls(form As Form)
        ' Apply resources to the form itself
        ApplyResourceToForm(form)
        
        ' Apply resources to controls within the form
        ApplyResourcesToControls(form, form.Name)
    End Sub

    ''' <summary>
    ''' Apply resources to the form itself.
    ''' </summary>
    ''' <param name="form">The form to localize.</param>
    Private Shared Sub ApplyResourceToForm(form As Form)
        ' Apply resources to the form using its name as the base resource key
        ApplyResourceToObject(form, form.Name)
    End Sub

    ''' <summary>
    ''' Apply resources to all controls within the specified parent control.
    ''' </summary>
    ''' <param name="parent">The parent control containing child controls to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourcesToControls(parent As Control, parentFormName As String)
        ' Iterate through each control within the parent control
        For Each ctrl As Control In parent.Controls
            ' Apply resources to the current control
            ApplyResourceToControl(ctrl, parentFormName)
            
            ' Recursively apply resources to child controls if the current control has children
            If ctrl.HasChildren Then
                ApplyResourcesToControls(ctrl, parentFormName)
            End If
        Next

        ' If the parent control is a form, also apply resources to MenuStrips within the form
        If TypeOf parent Is Form Then
            Dim form As Form = CType(parent, Form)
            For Each ctrl As Control In form.Controls
                If TypeOf ctrl Is MenuStrip Then
                    ApplyResourcesToMenuStrip(CType(ctrl, MenuStrip), parentFormName)
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Apply resources to a MenuStrip and its items.
    ''' </summary>
    ''' <param name="menuStrip">The MenuStrip to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        ' Iterate through each ToolStripMenuItem within the MenuStrip
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ' Apply resources to the current ToolStripMenuItem and its sub-items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    ''' <summary>
    ''' Apply resources to a ToolStripMenuItem and its sub-items.
    ''' </summary>
    ''' <param name="menuItem">The ToolStripMenuItem to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        ' Construct the resource key for the ToolStripMenuItem's text
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"

        Try
            ' Get the localized resource string for the constructed key
            Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
            
            ' If the resource value is not null, set the text of the ToolStripMenuItem
            If resourceValue IsNot Nothing Then
                menuItem.Text = resourceValue
            End If
        Catch ex As Exception
            ' Log an error message if an exception occurs while setting the text
            Debug.WriteLine($"Error setting text for menu item {menuItem.Name}: {ex.Message}")
        End Try

        ' Recursively apply resources to sub-items of the ToolStripMenuItem
        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem), parentFormName)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Apply resources to a control.
    ''' </summary>
    ''' <param name="control">The control to localize.</param>
    ''' <param
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourceToControl(control As Control, parentFormName As String)
        ' Apply resources to the control using its name as part of the resource key
        ApplyResourceToObject(control, $"{parentFormName}.{control.Name}")

        ' If the control is a ComboBox, apply resources to its items as well
        If TypeOf control Is ComboBox Then
            ApplyResourcesToComboBoxItems(CType(control, ComboBox), parentFormName)
        End If

        ' If the control is a GroupBox, apply resources to its child controls
        If TypeOf control Is GroupBox Then
            ApplyResourcesToGroupBox(CType(control, GroupBox), parentFormName)
        End If
    End Sub

    ''' <summary>
    ''' Apply resources to a GroupBox and its child controls.
    ''' </summary>
    ''' <param name="groupBox">The GroupBox to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
    Private Shared Sub ApplyResourcesToGroupBox(groupBox As GroupBox, parentFormName As String)
        ' Apply resources to the GroupBox itself
        ApplyResourceToControl(groupBox, parentFormName)

        ' Apply resources to controls within the GroupBox
        ApplyResourcesToControls(groupBox, parentFormName)
    End Sub

    ''' <summary>
    ''' Apply resources to the items of a ComboBox.
    ''' </summary>
    ''' <param name="comboBox">The ComboBox to localize.</param>
    ''' <param name="parentFormName">The name of the parent form.</param>
      Private Shared Sub ApplyResourcesToComboBoxItems(comboBox As ComboBox, parentFormName As String)
        ' Localize the Text property of the ComboBox
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Try
            Dim textResourceValue As String = GetLocalizedResourceByKey(textResourceKey)
            If textResourceValue IsNot Nothing Then
                comboBox.Text = textResourceValue
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error setting text for ComboBox {comboBox.Name}: {ex.Message}")
        End Try

        ' Localize the items of the ComboBox
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Try
                Dim resourceValue As String = GetLocalizedResourceByKey(resourceKey)
                If resourceValue IsNot Nothing Then
                    comboBox.Items.Add(resourceValue)
                Else
                    comboBox.Items.Add(items(i)) ' Add original item if no resource is found
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error setting item {i} for ComboBox {comboBox.Name}: {ex.Message}")
                comboBox.Items.Add(items(i)) ' Add original item if an error occurs
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Retrieve a localized resource string by its key.
    ''' </summary>
    ''' <param name="resourceKey">The resource key.</param>
    ''' <param name="values">Optional values to format the resource string.</param>
    ''' <returns>The localized resource string.</returns>
    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        ' Check if the resource key is already in the cache
        If ResourceCache.ContainsKey(resourceKey) Then
            ' If so, return the cached value
            Return If(values.Length > 0, String.Format(ResourceCache(resourceKey), values), ResourceCache(resourceKey))
        End If

        Try
            ' Attempt to retrieve the resource string from the resource manager
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                ' Cache the resource value for future use
                ResourceCache(resourceKey) = localizedResourceString
                ' Return the localized resource string
                Return If(values.Length > 0, String.Format(localizedResourceString, values), localizedResourceString)
            Else
                ' If the resource string is not found, return a placeholder
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            ' Log an error message if an exception occurs
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            ' Return an error message
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    ''' <summary>
    ''' Set the current UI culture.
    ''' </summary>
    ''' <param name="cultureName">The name of the culture to set.</param>
    Public Shared Sub SetCulture(cultureName As String)
        ' Set the current UI culture to the specified culture
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
        ' Clear the resource cache to ensure fresh localization data
        ResourceCache.Clear()
    End Sub
End Class

========================
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared ResourceCache As Dictionary(Of String, Dictionary(Of String, String))

    ' Initialize resource manager and cache
    Private Shared Sub InitializeResourceManager(resourceBaseName As String)
        ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
        ResourceCache = New Dictionary(Of String, Dictionary(Of String, String))
    End Sub

    ' Apply resources to a form and its controls
    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String)
        ' Initialize the resource manager and cache
        InitializeResourceManager(resourceBaseName)
        ' Apply resources to form and its controls
        ApplyResourcesRecursively(form, form.Name)
    End Sub

    ' Apply resources recursively to controls
    Private Shared Sub ApplyResourcesRecursively(parentControl As Control, parentFormName As String)
        ' Apply resources to parent control
        ApplyResourcesToControl(parentControl, parentFormName)

        ' Recursively apply resources to child controls
        For Each childControl As Control In parentControl.Controls
            ApplyResourcesRecursively(childControl, parentFormName)
        Next
    End Sub

    ' Apply resources to a control
    Private Shared Sub ApplyResourcesToControl(control As Control, parentFormName As String)
        ' Apply resources to control properties
        ApplyResourceToProperty(control, "Text", parentFormName)
        ApplyResourceToProperty(control, "Location", parentFormName)
        ApplyResourceToProperty(control, "ToolTipText", parentFormName)
        ' Add more properties as needed

        ' Handle specific control types
        If TypeOf control Is MenuStrip Then
            ApplyResourcesToMenuStrip(CType(control, MenuStrip), parentFormName)
        ElseIf TypeOf control Is ComboBox Then
            ApplyResourcesToComboBox(CType(control, ComboBox), parentFormName)
        ElseIf TypeOf control Is GroupBox Then
            ApplyResourcesToGroupBox(CType(control, GroupBox), parentFormName)
        End If
    End Sub

    ' Apply resources to a MenuStrip control
    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip, parentFormName As String)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem, parentFormName)
        Next
    End Sub

    ' Apply resources to a ToolStripMenuItem control
    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem, parentFormName As String)
        Dim resourceKey As String = $"{parentFormName}.{menuItem.Name}.Text"
        Dim resourceValue As String = GetCachedResourceValue(resourceKey)
        If resourceValue IsNot Nothing Then
            menuItem.Text = resourceValue
        End If
    End Sub

    ' Apply resources to a ComboBox control
    Private Shared Sub ApplyResourcesToComboBox(comboBox As ComboBox, parentFormName As String)
        ' Localize the Text property of the ComboBox
        Dim textResourceKey As String = $"{parentFormName}.{comboBox.Name}.Text"
        Dim textResourceValue As String = GetCachedResourceValue(textResourceKey)
        If textResourceValue IsNot Nothing Then
            comboBox.Text = textResourceValue
        End If

        ' Localize the items of the ComboBox
        Dim items = comboBox.Items.Cast(Of Object).ToArray()
        comboBox.Items.Clear()
        For i As Integer = 0 To items.Length - 1
            Dim resourceKey As String = $"{parentFormName}.{comboBox.Name}.Items[{i}]"
            Dim resourceValue As String = GetCachedResourceValue(resourceKey)
            If resourceValue IsNot Nothing Then
                comboBox.Items.Add(resourceValue)
            Else
                comboBox.Items.Add(items(i)) ' Add original item if no resource is found
            End If
        Next
    End Sub

    ' Apply resources to a GroupBox control
    Private Shared Sub ApplyResourcesToGroupBox(groupBox As GroupBox, parentFormName As String)
        ' Localize the Text property of the GroupBox
        Dim textResourceKey As String = $"{parentFormName}.{groupBox.Name}.Text"
        Dim textResourceValue As String = GetCachedResourceValue(textResourceKey)
        If textResourceValue IsNot Nothing Then
            groupBox.Text = textResourceValue
        End If

        ' Apply resources to controls within the GroupBox recursively
        For Each ctrl As Control In groupBox.Controls
            ApplyResourcesToControl(ctrl, $"{parentFormName}.{groupBox.Name}")
        Next
    End Sub

    ' Apply resource to a control property
    Private Shared Sub ApplyResourceToProperty(control As Control, propertyName As String, parentFormName As String)
        ' Generate resource key
        Dim resourceKey As String = $"{parentFormName}.{control.Name}.{propertyName}"
        Dim resourceValue As String = GetCachedResourceValue(resourceKey)
        If resourceValue IsNot Nothing Then
            ' Set property value
            Dim prop As Reflection.PropertyInfo = control.GetType().GetProperty(propertyName)
            If prop IsNot Nothing AndAlso prop.CanWrite Then
                prop.SetValue(control, Convert.ChangeType(resourceValue, prop.PropertyType), Nothing)
            End If
        End If
    End Sub

    ' Get resource value from cache or resource manager
    Private Shared Function GetCachedResourceValue(resourceKey As String) As String
        If ResourceCache.ContainsKey(resourceKey) Then
            Return ResourceCache(resourceKey)
        Else
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                ResourceCache.Add(resourceKey, resourceValue)
                Return resourceValue
            Catch ex As Exception
                Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
                Return Nothing
            End Try
        End If
    End Function

    ' Get localized resource by key with optional format values
    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            Dim localizedResourceString As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
            If localizedResourceString IsNot Nothing Then
                Return String.Format(localizedResourceString, values)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function

    ' Set the current UI culture
    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub
End Class
===============Usung cascade cache
Imports System.Globalization
Imports System.Threading
Imports System.Windows.Forms
Imports System.Resources
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Diagnostics

Public Class ResourceHelper
    Private Shared ResourceManager As ResourceManager
    Private Shared ResourceCache As Dictionary(Of String, Dictionary(Of String, String))

    ' Initialize resource manager and cache
    Private Shared Sub InitializeResourceManager(resourceBaseName As String, cultureName As String)
        If ResourceManager Is Nothing OrElse Thread.CurrentThread.CurrentUICulture.Name <> cultureName Then
            ResourceManager = New ResourceManager(resourceBaseName, GetType(ResourceHelper).Assembly)
            Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
        End If

        If ResourceCache Is Nothing Then
            ResourceCache = New Dictionary(Of String, Dictionary(Of String, String))()
        End If
    End Sub

    ' Apply resources to a form and its controls
    Public Shared Sub ApplyResources(form As Form, resourceBaseName As String, cultureName As String)
        InitializeResourceManager(resourceBaseName, cultureName)
        ApplyResourceToForm(form)
        ApplyResourcesRecursively(form)
    End Sub

    ' Apply resources recursively to controls
    Private Shared Sub ApplyResourcesRecursively(parentControl As Control)
        ApplyResourceToControl(parentControl)
        For Each childControl As Control In parentControl.Controls
            ApplyResourcesRecursively(childControl)
        Next
    End Sub

    ' Apply resources to a control
    Private Shared Sub ApplyResourceToControl(control As Control)
        ApplyResourceToProperty(control, "Text")
        ApplyResourceToProperty(control, "Location")
        ApplyResourceToProperty(control, "ToolTipText")

        If TypeOf control Is MenuStrip Then
            ApplyResourcesToMenuStrip(CType(control, MenuStrip))
        End If
    End Sub

    ' Apply resources to a MenuStrip control
    Private Shared Sub ApplyResourcesToMenuStrip(menuStrip As MenuStrip)
        For Each menuItem As ToolStripMenuItem In menuStrip.Items
            ApplyResourceToMenuItem(menuItem)
        Next
    End Sub

    ' Apply resources to a ToolStripMenuItem control
    Private Shared Sub ApplyResourceToMenuItem(menuItem As ToolStripMenuItem)
        ApplyResourceToProperty(menuItem, "Text")

        For Each dropDownItem As ToolStripItem In menuItem.DropDownItems
            If TypeOf dropDownItem Is ToolStripMenuItem Then
                ApplyResourceToMenuItem(CType(dropDownItem, ToolStripMenuItem))
            End If
        Next
    End Sub

    ' Apply resource to a control property
    Private Shared Sub ApplyResourceToProperty(control As Control, propertyName As String)
        Dim resourceKey As String = $"{control.Name}.{propertyName}"
        Dim resourceValue As String = GetCachedResourceValue(resourceKey)
        If resourceValue IsNot Nothing Then
            Dim prop As Reflection.PropertyInfo = control.GetType().GetProperty(propertyName)
            If prop IsNot Nothing AndAlso prop.CanWrite Then
                Dim convertedValue As Object = ConvertToType(resourceValue, prop.PropertyType)
                If convertedValue IsNot Nothing Then
                    prop.SetValue(control, convertedValue, Nothing)
                End If
            End If
        End If
    End Sub

    ' Get resource value from cache or resource manager
    Private Shared Function GetCachedResourceValue(resourceKey As String) As String
        Dim currentCulture As String = Thread.CurrentThread.CurrentUICulture.Name

        If Not ResourceCache.ContainsKey(currentCulture) Then
            ResourceCache(currentCulture) = New Dictionary(Of String, String)()
        End If

        Dim cultureCache As Dictionary(Of String, String) = ResourceCache(currentCulture)

        If cultureCache.ContainsKey(resourceKey) Then
            Return cultureCache(resourceKey)
        Else
            Try
                Dim resourceValue As String = ResourceManager.GetString(resourceKey, Thread.CurrentThread.CurrentUICulture)
                If resourceValue IsNot Nothing Then
                    cultureCache(resourceKey) = resourceValue
                End If
                Return resourceValue
            Catch ex As Exception
                Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
                Return Nothing
            End Try
        End If
    End Function

    ' Convert string value to target property type
    Private Shared Function ConvertToType(value As String, targetType As Type) As Object
        Try
            If targetType Is GetType(String) Then
                Return value
            ElseIf targetType Is GetType(Integer) Then
                Return Integer.Parse(value)
            ElseIf targetType Is GetType(Boolean) Then
                Return Boolean.Parse(value)
            ElseIf targetType Is GetType(Point) Then
                Dim parts() As String = value.Split(","c)
                Return New Point(Integer.Parse(parts(0)), Integer.Parse(parts(1)))
            Else
                Return Convert.ChangeType(value, targetType)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error converting value '{value}' to type '{targetType.Name}': {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' Apply resource to form name
    Private Shared Sub ApplyResourceToForm(form As Form)
        Dim resourceKey As String = $"{form.Name}.Text"
        Dim resourceValue As String = GetCachedResourceValue(resourceKey)
        If resourceValue IsNot Nothing Then
            form.Text = resourceValue
        End If
    End Sub

    ' Set the culture for resource localization
    Public Shared Sub SetCulture(cultureName As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
    End Sub

    ' Get localized resource by key with optional formatting values
    Public Shared Function GetLocalizedResourceByKey(resourceKey As String, ParamArray values As Object()) As String
        Try
            Dim localizedResourceString As String = GetCachedResourceValue(resourceKey)
            If localizedResourceString IsNot Nothing Then
                Return String.Format(localizedResourceString, values)
            Else
                Return $"[Missing Resource: {resourceKey}]"
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error retrieving resource {resourceKey}: {ex.Message}")
            Return $"[Error: {resourceKey}]"
        End Try
    End Function
End Class


