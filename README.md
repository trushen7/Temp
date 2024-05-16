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

