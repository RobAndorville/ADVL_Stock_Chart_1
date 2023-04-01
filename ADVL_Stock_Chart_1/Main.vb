
'==============================================================================================================================================================================================
'
'Copyright 2018 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Imports System.Windows.Forms.DataVisualization
Imports System.Security.Permissions
Imports System.ComponentModel
Imports System.Windows.Forms.DataVisualization.Charting

<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class Main
    'The ADVL_Stock_Chart application produces charts of stock prices.

#Region " Coding Notes - Notes on the code used in this class." '==============================================================================================================================

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'ADVL_Utilities_Library_1.dll
    'To add the reference, press Project \ Add Reference... 
    '  Select the Browse option then press the Browse button
    '  Find the ADVL_Utilities_Library_1.dll file (it should be located in the directory ...\Projects\ADVL_Utilities_Library_1\ADVL_Utilities_Library_1\bin\Debug\)
    '  Press the Add button. Press the OK button.
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'ADD THE SERVICE REFERENCE: ===================================================================================================
    'A service reference to the Message Service must be added to the source code before this service can be used.
    'This is used to connect to the Application Network.

    'Adding the service reference to a project that includes the WcfMsgServiceLib project: -----------------------------------------
    'Project \ Add Service Reference
    'Press the Discover button.
    'Expand the items in the Services window and select IMsgService.
    'Press OK.
    '------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------
    'Adding the service reference to other projects that dont include the WcfMsgServiceLib project: -------------------------------
    'Run the ADVL_Application_Network_1 application to start the Application Network message service.
    'In Microsoft Visual Studio select: Project \ Add Service Reference
    'Enter the address: http://localhost:8734/ADVLService
    'Press the Go button.
    'MsgService is found.
    'Press OK to add ServiceReference1 to the project.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE MsgServiceCallback CODE: =============================================================================================
    'This is used to connect to the Application Network.
    'In Microsoft Visual Studio select: Project \ Add Class
    'MsgServiceCallback.vb
    'Add the following code to the class:
    'Imports System.ServiceModel
    'Public Class MsgServiceCallback
    '    Implements ServiceReference1.IMsgServiceCallback
    '    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    '        'A message has been received.
    '        'Set the InstrReceived property value to the message (usually in XMessage format). This will also apply the instructions in the XMessage.
    '        Main.InstrReceived = message
    '    End Sub
    'End Class
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'DEBUGGING TIPS:
    '1. If an application based on the Application Template does not initially run correctly,
    '    check that the copied methods, such as Main_Load, have the correct Handles statement.
    '    For example: the Main_Load method should have the following declaration: Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
    '      It will not run when the application loads, with this declaration:      Private Sub Main_Load(sender As Object, e As EventArgs)
    '    For example: the Main_FormClosing method should have the following declaration: Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    '      It will not run when the application closes, with this declaration:     Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs)
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE Timer1 Control to the Main Form: =====================================================================================
    'Select the Main.vb [Design] tab.
    'Press Toolbox \ Compnents \ Times and add Timer1 to the Main form.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'EDIT THE DefaultAppProperties() CODE: ========================================================================================
    'This sets the Application properties that are stored in the Application_Info_ADVL_2.xml settings file.
    'The following properties need to be updated:
    '  ApplicationInfo.Name
    '  ApplicationInfo.Description
    '  ApplicationInfo.CreationDate
    '  ApplicationInfo.Author
    '  ApplicationInfo.Copyright
    '  ApplicationInfo.Trademarks
    '  ApplicationInfo.License
    '  ApplicationInfo.SourceCode          (Optional - Preliminary implemetation coded.)
    '  ApplicationInfo.ModificationSummary (Optional - Preliminary implemetation coded.)
    '  ApplicationInfo.Libraries           (Optional - Preliminary implemetation coded.)
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE Application Icon: ====================================================================================================
    'Double-click My Project in the Solution Explorer window to open the project tab.
    'In the Application section press the Icon box and selct Browse.
    'Select an application icon.
    'This icon can also be selected for the Main form icon by editing the properties of this form.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'EDIT THE Application Info Text: ==============================================================================================
    'The Application Info Text is used to label the appllication icon in the Application Network tree view.
    'This is edited in the SendApplicationInfo() method of the Main form.
    'Edit the line of code: Dim text As New XElement("Text", "Application Template").
    'Replace the default text "Application Template" with the required text.
    'Note that this text can be updated at any time and when the updated executable is run, it will update the Application Network tree view the next time it is connected.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'Calling JavaScript from VB.NET:
    'The following Imports statement and permissions are required for the Main form:
    'Imports System.Security.Permissions
    '<PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
    '<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
    'NOTE: the line continuation characters (_) will disappear form the code view after they have been typed!
    '------------------------------------------------------------------------------------------------------------------------------
    'Calling VB.NET from JavaScript
    'Add the following line to the Main.Load method:
    '  Me.WebBrowser1.ObjectForScripting = Me
    '------------------------------------------------------------------------------------------------------------------------------



#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables and class objects used in this form and this application." '===============================================================================

    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.

    'Declare Forms used by the application:
    Public WithEvents DesignQuery As frmDesignQuery
    Public WithEvents ViewDatabaseData As frmViewDatabaseData
    Public WithEvents WebPageList As frmWebPageList
    Public WithEvents ProjectArchive As frmArchive 'Form used to view the files in a Project archive
    Public WithEvents SettingsArchive As frmArchive 'Form used to view the files in a Settings archive
    Public WithEvents DataArchive As frmArchive 'Form used to view the files in a Data archive
    Public WithEvents SystemArchive As frmArchive 'Form used to view the files in a System archive

    Public WithEvents NewHtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    Public WithEvents ZoomChart As frmZoomChart 'Used to zoom the chart view.

    'Declare objects used to connect to the Communication Network:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientProNetName As String = "" 'The name of the client Project Network requesting service.
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    Public OnCompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public EndInstruction As String = "Stop" 'Another method of specifying the last instruction. This is processed in the EndOfSequence section of XMsg.Instructions.

    Public ConnectionName As String = "" 'The name of the connection used to connect this application to the ComNet.

    Public ProNetName As String = "" 'The name of the Project Network
    Public ProNetPath As String = "" 'The path of the Project Network

    Public AdvlNetworkAppPath As String = "" 'The application path of the ADVL Network application (ComNet). This is where the "Application.Lock" file will be while ComNet is running
    Public AdvlNetworkExePath As String = "" 'The executable path of the ADVL Network.

    'Variable for local processing of an XMessage:
    Public WithEvents XMsgLocal As New ADVL_Utilities_Library_1.XMessage
    Dim XDocLocal As New System.Xml.XmlDocument
    Public StatusLocal As New System.Collections.Specialized.StringCollection

    'Main.Load variables:
    Dim ProjectSelected As Boolean = False 'If True, a project has been selected using Command Arguments. Used in Main.Load.
    Dim StartupConnectionName As String = "" 'If not "" the application will be connected to the ComNet using this connection name in  Main.Load.

    'The following variables are used to run JavaScript in Web Pages loaded into the Document View: -------------------
    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence
    'To run an XSequence:
    '  XSeq.RunXSequence(xDoc, Status) 'ImportStatus in Import
    '    Handle events:
    '      XSeq.ErrorMsg
    '      XSeq.Instruction(Info, Locn)

    Private XStatus As New System.Collections.Specialized.StringCollection

    'Variables used to restore Item values on a web page.
    Private FormName As String
    Private ItemName As String
    Private SelectId As String

    'StartProject variables:
    Private StartProject_AppName As String  'The application name
    Private StartProject_ConnName As String 'The connection name
    Private StartProject_ProjID As String   'The project ID

    Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Chart Settings tab

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

    Private WithEvents bgwComCheck As New System.ComponentModel.BackgroundWorker

    Public WithEvents bgwSendMessage As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service.
    Dim SendMessageParams As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message

    'Alternative SendMessage background worker - needed to send a message while instructions are being processed.
    Public WithEvents bgwSendMessageAlt As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service - alternative backgound worker.
    Dim SendMessageParamsAlt As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message - for the alternative background worker.


    Public WithEvents ChartInfo As New ChartInfo 'Stores information about the chart. Contains methods to Save, Load and Clear the chart.

    'Variables used to process Stock Chart instructions:
    Dim SeriesInfoName As String = ""
    Dim AreaInfoName As String = ""
    Dim TitleName As String = ""
    Dim ChartFontName As String = "Arial"
    Dim ChartFontStyle As FontStyle
    Dim ChartFontSize As Single = 12
    Dim SeriesName As String = ""
    Dim AreaName As String = ""

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _connectionHashcode As Integer 'The Application Network connection hashcode. This is used to identify a connection in the Application Netowrk when reconnecting.
    Property ConnectionHashcode As Integer
        Get
            Return _connectionHashcode
        End Get
        Set(value As Integer)
            _connectionHashcode = value
        End Set
    End Property


    Private _connectedToComNet As Boolean = False  'True if the application is connected to the Communication Network (Message Service).
    Property ConnectedToComNet As Boolean
        Get
            Return _connectedToComNet
        End Get
        Set(value As Boolean)
            _connectedToComNet = value
        End Set
    End Property


    Private _instrReceived As String = "" 'Contains Instructions received from the Application Network message service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value
                ProcessInstructions(_instrReceived)
            End If
        End Set
    End Property

    Private Sub ProcessInstructions(ByVal Instructions As String)
        'Process the XMessage instructions.

        Dim MsgType As String
        If Instructions.StartsWith("<XMsg>") Then
            MsgType = "XMsg"
            If ShowXMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        ElseIf Instructions.StartsWith("<XSys>") Then
            MsgType = "XSys"
            If ShowSysMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("System Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        Else
            MsgType = "Unknown"
        End If

        'If ShowXMessages Then
        '    'Add the message header to the XMessages window:
        '    Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
        'End If

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If MsgType = "XMsg" Or MsgType = "XSys" Then 'This is an XMessage or XSystem set of instructions.
                Try
                    'Inititalise the reply message:
                    ClientProNetName = ""
                    ClientConnName = ""
                    ClientAppName = ""
                    xlocns.Clear() 'Clear the list of locations in the reply message.
                    Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                    MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                'xmessage = New XElement("XMsg")
                xmessage = New XElement(MsgType)
                xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                    'Run the received message:
                    Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                    XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                'If ShowXMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                XMsg.Run(XDoc, Status)
                Catch ex As Exception
                    Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                End Try

                'XMessage has been run.
                'Reply to this message:
                'Add the message reply to the XMessages window:
                'Complete the MessageXDoc:
                xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the last location reply instructions to the message.
                MessageXDoc.Add(xmessage)
                MessageText = MessageXDoc.ToString

                If ClientConnName = "" Then
                'No client to send a message to - process the message locally.
                'If ShowXMessages Then
                '    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                ProcessLocalInstructions(MessageText)
                Else
                'If ShowXMessages Then
                '    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Send Message on a new thread:
                SendMessageParams.ProjectNetworkName = ClientProNetName
                    SendMessageParams.ConnectionName = ClientConnName
                    SendMessageParams.Message = MessageText
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    End If
                End If

            Else 'This is not an XMessage!
                If Instructions.StartsWith("<XMsgBlk>") Then 'This is an XMessageBlock.
                'Process the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                If ShowXMessages Then
                    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Process the XMessageBlock:
                Dim XMsgBlkLocn As String
                XMsgBlkLocn = XDoc.GetElementsByTagName("ClientLocn")(0).InnerText
                Select Case XMsgBlkLocn
                    Case "DisplayChart"
                        Dim XData As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo")
                        Dim ChartXDoc As New Xml.Linq.XDocument
                        Try
                            ChartXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XData(0).InnerXml)
                            ChartInfo.LoadXml(ChartXDoc, Chart1)
                            txtChartFileName.Text = ""
                            txtSeriesName.Text = ChartXDoc.<ChartSettings>.<SeriesCollection>.<Series>.<Name>.Value
                            UpdateInputDataTabSettings()
                            UpdateTitlesTabSettings()
                            UpdateAreasTabSettings() 'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                            UpdateSeriesTabSettings()
                            DrawStockChart()

                        Catch ex As Exception
                            Message.Add(ex.Message & vbCrLf)
                        End Try

                    Case Else
                        Message.AddWarning("Unknown XInfo Message location: " & XMsgBlkLocn & vbCrLf)
                End Select
            Else
                Message.XAddText("The message is not an XMessage or XMessageBlock: " & vbCrLf & Instructions & vbCrLf & vbCrLf, "Normal")
            End If
        End If
    End Sub

    Private Sub ProcessLocalInstructions(ByVal Instructions As String)
        'Process the XMessage instructions locally.

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If Instructions.StartsWith("<XMsg>") Or Instructions.StartsWith("<XSys>") Then 'This is an XMessage set of instructions.
                'Run the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDocLocal.LoadXml(XmlHeader & vbCrLf & Instructions)
                XMsgLocal.Run(XDocLocal, StatusLocal)
            Else 'This is not an XMessage!
                Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private _showXMessages As Boolean = True 'If True, XMessages that are sent or received will be shown in the Messages window.
    Property ShowXMessages As Boolean
        Get
            Return _showXMessages
        End Get
        Set(value As Boolean)
            _showXMessages = value
        End Set
    End Property

    Private _showSysMessages As Boolean = True 'If True, System messages that are sent or received will be shown in the messages window.
    Property ShowSysMessages As Boolean
        Get
            Return _showSysMessages
        End Get
        Set(value As Boolean)
            _showSysMessages = value
        End Set
    End Property

    Private _closedFormNo As Integer 'Temporarily holds the number of the form that is being closed. 
    Property ClosedFormNo As Integer
        Get
            Return _closedFormNo
        End Get
        Set(value As Integer)
            _closedFormNo = value
        End Set
    End Property

    Private _inputDatabaseDirectory As String = "" 'The directory of the Input Database. When the Find database button is pressed, the Open File Dialog will open in this directory.
    Property InputDatabaseDirectory As String
        Get
            Return _inputDatabaseDirectory
        End Get
        Set(value As String)
            _inputDatabaseDirectory = value
        End Set
    End Property


    Private _chartWindow As String = "Preview" '(Preview or New Window) Chart can be drawn in the Preview window or a New Window.
    Property ChartWindow As String
        Get
            Return _chartWindow
        End Get
        Set(value As String)
            _chartWindow = value
            If _chartWindow = "Preview" Then
                rbPreviewChart.Checked = True
            ElseIf _chartWindow = "New Window" Then
                rbNewWindowChart.Checked = True
            End If
        End Set
    End Property

    Private _workflowFileName As String = "" 'The file name of the html document displayed in the Workflow tab.
    Public Property WorkflowFileName As String
        Get
            Return _workflowFileName
        End Get
        Set(value As String)
            _workflowFileName = value
        End Set
    End Property

    Private _selectedChart As DataVisualization.Charting.Chart 'The selected chart.
    Property SelectedChart As DataVisualization.Charting.Chart
        Get
            Return _selectedChart
        End Get
        Set(value As DataVisualization.Charting.Chart)
            _selectedChart = value
        End Set
    End Property

    Private _selectedChartInfo As ChartInfo 'The ChartInfo correspoinding to the selected Chart.
    Property SelectedChartInfo As ChartInfo
        Get
            Return _selectedChartInfo
        End Get
        Set(value As ChartInfo)
            _selectedChartInfo = value
        End Set
    End Property

    Private _selectedChartNo As Integer = -1 'The selected chart number. If -1 then Chart1 on the Main form has been selected. If 0 or greater, the corresponding Chart on the ChartList() has been selected.
    Property SelectedChartNo As Integer
        Get
            Return _selectedChartNo
        End Get
        Set(value As Integer)
            _selectedChartNo = value
        End Set
    End Property


#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML Files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <AdvlNetworkAppPath><%= AdvlNetworkAppPath %></AdvlNetworkAppPath>
                               <AdvlNetworkExePath><%= AdvlNetworkExePath %></AdvlNetworkExePath>
                               <ShowXMessages><%= ShowXMessages %></ShowXMessages>
                               <ShowSysMessages><%= ShowSysMessages %></ShowSysMessages>
                               <WorkFlowFileName><%= WorkflowFileName %></WorkFlowFileName>
                               <!---->
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <StyleSplitterDistance><%= SplitContainer1.SplitterDistance %></StyleSplitterDistance>
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            If Settings.<FormSettings>.<AdvlNetworkAppPath>.Value <> Nothing Then AdvlNetworkAppPath = Settings.<FormSettings>.<AdvlNetworkAppPath>.Value
            If Settings.<FormSettings>.<AdvlNetworkExePath>.Value <> Nothing Then AdvlNetworkExePath = Settings.<FormSettings>.<AdvlNetworkExePath>.Value

            If Settings.<FormSettings>.<ShowXMessages>.Value <> Nothing Then ShowXMessages = Settings.<FormSettings>.<ShowXMessages>.Value
            If Settings.<FormSettings>.<ShowSysMessages>.Value <> Nothing Then ShowSysMessages = Settings.<FormSettings>.<ShowSysMessages>.Value

            If Settings.<FormSettings>.<WorkFlowFileName>.Value <> Nothing Then WorkflowFileName = Settings.<FormSettings>.<WorkFlowFileName>.Value

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value

            If Settings.<FormSettings>.<StyleSplitterDistance>.Value <> Nothing Then SplitContainer1.SplitterDistance = Settings.<FormSettings>.<StyleSplitterDistance>.Value
            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
        End If
    End Sub

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info_ADVL_2.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties:
            ApplicationInfo.WriteFile() 'Write the file now. The file information may be used by other applications.
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        ApplicationInfo.Name = "ADVL_Stock_Chart_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        ApplicationInfo.Description = "The ADVL_Stock_Chart application produces charts of stock prices."
        ApplicationInfo.CreationDate = "18-Dec-2018 12:00:00"

        'Author -----------------------------------------------------------------------------------------------------------
        'Change this to show your Name, Description and Contact information.
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville™ software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2018"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)
        Dim Trademark3 As New ADVL_Utilities_Library_1.Trademark
        Trademark3.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark3.Text = "AL-M7"
        Trademark3.Registered = False
        Trademark3.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark3)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2018"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code is used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville™ software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville™ software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville™ software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville™ software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville™ software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville™ software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville™ software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub

    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save the project settings in an XML file.

        'Add any Project Settings to be saved into the settingsData XDocument.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Stock_Chart_1 application.-->
                           <ProjectSettings>
                               <!--Plot Settings:-->
                               <ChartWindow><%= ChartWindow %></ChartWindow>
                               <AutoDraw><%= chkAutoDraw.Checked %></AutoDraw>
                               <ChartFileName><%= txtChartFileName.Text %></ChartFileName>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)



        'OLD CODE:
        'Add any Project Settings to be saved into the settingsData XDocument.
        'Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
        '                   <!---->
        '                   <!--Project settings for ADVL_Stock_Chart_1 application.-->
        '                   <ProjectSettings>
        '                       <!--File Name:-->
        '                       <FileName><%= StockChart.FileName %></FileName>
        '                       <!--Plot Settings:-->
        '                       <ChartWindow><%= ChartWindow %></ChartWindow>
        '                       <AutoDraw><%= chkAutoDraw.Checked %></AutoDraw>
        '                       <!--Input Data:-->
        '                       <InputDataType><%= StockChart.InputDataType %></InputDataType>
        '                       <InputDatabasePath><%= StockChart.InputDatabasePath %></InputDatabasePath>
        '                       <InputDatabaseDirectory><%= InputDatabaseDirectory %></InputDatabaseDirectory>
        '                       <InputDataDescription><%= StockChart.InputDataDescr %></InputDataDescription>
        '                       <InputQuery><%= StockChart.InputQuery %></InputQuery>
        '                       <!--Names of the Fields plotted on the chart:-->
        '                       <YValuesHighFieldName><%= StockChart.YValuesHighFieldName %></YValuesHighFieldName>
        '                       <YValuesLowFieldName><%= StockChart.YValuesLowFieldName %></YValuesLowFieldName>
        '                       <YValuesOpenFieldName><%= StockChart.YValuesOpenFieldName %></YValuesOpenFieldName>
        '                       <YValuesCloseFieldName><%= StockChart.YValuesCloseFieldName %></YValuesCloseFieldName>
        '                       <XValuesFieldName><%= StockChart.XValuesFieldName %></XValuesFieldName>
        '                       <!--Chart Settings:-->
        '                       <SeriesName><%= StockChart.SeriesName %></SeriesName>
        '                       <LabelValueType><%= StockChart.LabelValueType %></LabelValueType>
        '                       <MaxPixelPointWidth><%= StockChart.MaxPixelPointWidth %></MaxPixelPointWidth>
        '                       <MinPixelPointWidth><%= StockChart.MinPixelPointWidth %></MinPixelPointWidth>
        '                       <OpenCloseStyle><%= StockChart.OpenCloseStyle %></OpenCloseStyle>
        '                       <PixelPointDepth><%= StockChart.PixelPointDepth %></PixelPointDepth>
        '                       <PixelPointGapDepth><%= StockChart.PixelPointGapDepth %></PixelPointGapDepth>
        '                       <PixelPointWidth><%= StockChart.PixelPointWidth %></PixelPointWidth>
        '                       <PointWidth><%= StockChart.PointWidth %></PointWidth>
        '                       <ShowOpenClose><%= StockChart.ShowOpenClose %></ShowOpenClose>
        '                       <!--Stock Chart Label:-->
        '                       <ChartLabelName><%= StockChart.ChartLabel.Name %></ChartLabelName>
        '                       <ChartLabelFontName><%= StockChart.ChartLabel.FontName %></ChartLabelFontName>
        '                       <ChartLabelSize><%= StockChart.ChartLabel.Size %></ChartLabelSize>
        '                       <ChartLabelBold><%= StockChart.ChartLabel.Bold %></ChartLabelBold>
        '                       <ChartLabelItalic><%= StockChart.ChartLabel.Italic %></ChartLabelItalic>
        '                       <ChartLabelUnderline><%= StockChart.ChartLabel.Underline %></ChartLabelUnderline>
        '                       <ChartLabelStrikeout><%= StockChart.ChartLabel.Strikeout %></ChartLabelStrikeout>
        '                       <ChartLabelText><%= StockChart.ChartLabel.Text %></ChartLabelText>
        '                       <ChartLabelColor><%= StockChart.ChartLabel.Color.ToArgb.ToString %></ChartLabelColor>
        '                       <ChartLabelAlignment><%= StockChart.ChartLabel.Alignment.ToString %></ChartLabelAlignment>
        '                       <!--X Axis:-->
        '                       <XAxisTitleText><%= StockChart.XAxis.Title.Text %></XAxisTitleText>
        '                       <XAxisTitleFontName><%= StockChart.XAxis.Title.FontName %></XAxisTitleFontName>
        '                       <XAxisTitleColor><%= StockChart.XAxis.Title.Color %></XAxisTitleColor>
        '                       <XAxisTitleSize><%= StockChart.XAxis.Title.Size %></XAxisTitleSize>
        '                       <XAxisTitleBold><%= StockChart.XAxis.Title.Bold %></XAxisTitleBold>
        '                       <XAxisTitleItalic><%= StockChart.XAxis.Title.Italic %></XAxisTitleItalic>
        '                       <XAxisTitleUnderline><%= StockChart.XAxis.Title.Underline %></XAxisTitleUnderline>
        '                       <XAxisTitleStrikeout><%= StockChart.XAxis.Title.Strikeout %></XAxisTitleStrikeout>
        '                       <XAxisTitleAlignment><%= StockChart.XAxis.TitleAlignment %></XAxisTitleAlignment>
        '                       <XAxisAutoMinimum><%= StockChart.XAxis.AutoMinimum %></XAxisAutoMinimum>
        '                       <XAxisMinimum><%= StockChart.XAxis.Minimum %></XAxisMinimum>
        '                       <XAxisAutoMaximum><%= StockChart.XAxis.AutoMaximum %></XAxisAutoMaximum>
        '                       <XAxisMaximum><%= StockChart.XAxis.Maximum %></XAxisMaximum>
        '                       <XAxisAutoInterval><%= StockChart.XAxis.AutoInterval %></XAxisAutoInterval>
        '                       <XAxisInterval><%= StockChart.XAxis.Interval %></XAxisInterval>
        '                       <XAxisAutoMajorGridInterval><%= StockChart.XAxis.AutoMajorGridInterval %></XAxisAutoMajorGridInterval>
        '                       <XAxisMajorGridInterval><%= StockChart.XAxis.MajorGridInterval %></XAxisMajorGridInterval>
        '                       <!--Y Axis:-->
        '                       <YAxisTitleText><%= StockChart.YAxis.Title.Text %></YAxisTitleText>
        '                       <YAxisTitleFontName><%= StockChart.YAxis.Title.FontName %></YAxisTitleFontName>
        '                       <YAxisTitleColor><%= StockChart.YAxis.Title.Color %></YAxisTitleColor>
        '                       <YAxisTitleSize><%= StockChart.YAxis.Title.Size %></YAxisTitleSize>
        '                       <YAxisTitleBold><%= StockChart.YAxis.Title.Bold %></YAxisTitleBold>
        '                       <YAxisTitleItalic><%= StockChart.YAxis.Title.Italic %></YAxisTitleItalic>
        '                       <YAxisTitleUnderline><%= StockChart.YAxis.Title.Underline %></YAxisTitleUnderline>
        '                       <YAxisTitleStrikeout><%= StockChart.YAxis.Title.Strikeout %></YAxisTitleStrikeout>
        '                       <YAxisTitleAlignment><%= StockChart.YAxis.TitleAlignment %></YAxisTitleAlignment>
        '                       <YAxisAutoMinimum><%= StockChart.YAxis.AutoMinimum %></YAxisAutoMinimum>
        '                       <YAxisMinimum><%= StockChart.YAxis.Minimum %></YAxisMinimum>
        '                       <YAxisAutoMaximum><%= StockChart.YAxis.AutoMaximum %></YAxisAutoMaximum>
        '                       <YAxisMaximum><%= StockChart.YAxis.Maximum %></YAxisMaximum>
        '                       <YAxisAutoInterval><%= StockChart.YAxis.AutoInterval %></YAxisAutoInterval>
        '                       <YAxisInterval><%= StockChart.YAxis.Interval %></YAxisInterval>
        '                       <YAxisAutoMajorGridInterval><%= StockChart.YAxis.AutoMajorGridInterval %></YAxisAutoMajorGridInterval>
        '                       <YAxisMajorGridInterval><%= StockChart.YAxis.MajorGridInterval %></YAxisMajorGridInterval>
        '                   </ProjectSettings>

        ''   <ChartLabelColor><%= StockChart.ChartLabel.Color.ToString %></ChartLabelColor>

        'Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & ".xml"
        'Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreProjectSettings()
        'Restore the project settings from an XML document.

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                'Exit Sub
            Else
                ''Restore a Project Setting example:
                'If Settings.<ProjectSettings>.<Setting1>.Value = Nothing Then
                '    'Project setting not saved.
                '    'Setting1 = ""
                'Else
                '    'Setting1 = Settings.<ProjectSettings>.<Setting1>.Value
                'End If

                'Restore Plot Settings: ==============================================================================================
                If Settings.<ProjectSettings>.<ChartWindow>.Value <> Nothing Then ChartWindow = Settings.<ProjectSettings>.<ChartWindow>.Value
                If Settings.<ProjectSettings>.<AutoDraw>.Value <> Nothing Then chkAutoDraw.Checked = Settings.<ProjectSettings>.<AutoDraw>.Value
                If Settings.<ProjectSettings>.<ChartFileName>.Value <> Nothing Then txtChartFileName.Text = Settings.<ProjectSettings>.<ChartFileName>.Value
                '---------------------------------------------------------------------------------------------------------------------

                'Continue restoring saved settings.

            End If
        End If
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Loading the Main form.

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property

        ''Get the Application Version Information:
        'ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        'ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        'ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        'ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
                Exit Sub
            End If
        End If

        ReadApplicationInfo()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()

        'Restore Project information: -------------------------------------------------------
        Project.Application.Name = ApplicationInfo.Name

        'Set up Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form - This will show the App icon on top in the TaskBar.

        'Start showing messages here - Message system is set up.
        'Message.AddText("------------------- Starting Application: ADVL Application Template ----------------- " & vbCrLf, "Heading")
        Message.AddText("------------------- Starting Application: ADVL Stock Chart ----------------- " & vbCrLf, "Heading")
        'Message.AddText("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#.##") & " hours" & vbCrLf, "Normal")
        Dim TotalDuration As String = ApplicationUsage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                           ApplicationUsage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                           ApplicationUsage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                           ApplicationUsage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"
        Message.AddText("Application usage: Total duration = " & TotalDuration & vbCrLf, "Normal")

        'https://msdn.microsoft.com/en-us/library/z2d603cy(v=vs.80).aspx#Y550
        'Process any command line arguments:
        Try
            For Each s As String In My.Application.CommandLineArgs
                Message.Add("Command line argument: " & vbCrLf)
                Message.AddXml(s & vbCrLf & vbCrLf)
                InstrReceived = s
            Next
        Catch ex As Exception
            Message.AddWarning("Error processing command line arguments: " & ex.Message & vbCrLf)
        End Try

        If ProjectSelected = False Then
            'Read the Settings Location for the last project used:
            Project.ReadLastProjectInfo()
            'The Last_Project_Info.xml file contains:
            '  Project Name and Description. Settings Location Type and Settings Location Path.
            Message.Add("Last project details:" & vbCrLf)
            Message.Add("Project Type:  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project Path:  " & Project.Path & vbCrLf)

            'At this point read the application start arguments, if any.
            'The selected project may be changed here.

            'Check if the project is locked:
            If Project.ProjectLocked Then
                Message.AddWarning("The project is locked: " & Project.Name & vbCrLf)
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Press 'Yes' to unlock the project", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.Yes Then
                    Project.UnlockProject()
                    Message.AddWarning("The project has been unlocked: " & Project.Name & vbCrLf)
                    'Read the Project Information file: -------------------------------------------------
                    Message.Add("Reading project info." & vbCrLf)
                    Project.ReadProjectInfoFile()   'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                    Project.ReadParameters()
                    Project.ReadParentParameters()
                    If Project.ParentParameterExists("ProNetName") Then
                        Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetName = Project.Parameter("ProNetName").Value
                    Else
                        ProNetName = Project.GetParameter("ProNetName")
                    End If
                    If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                        Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetPath = Project.Parameter("ProNetPath").Value
                    Else
                        ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                    End If
                    Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                    Message.Show()
                Else
                    'Continue without any project selected.
                    Project.Name = ""
                    Project.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Project.Description = ""
                    Project.SettingsLocn.Path = ""
                    Project.DataLocn.Path = ""
                End If
            Else
                'Read the Project Information file: -------------------------------------------------
                Message.Add("Reading project info." & vbCrLf)
                Project.ReadProjectInfoFile()  'Read the file in the SettingsLocation: ADVL_Project_Info.xml

                Project.ReadParameters()
                Project.ReadParentParameters()
                If Project.ParentParameterExists("ProNetName") Then
                    Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetName = Project.Parameter("ProNetName").Value
                Else
                    ProNetName = Project.GetParameter("ProNetName")
                End If
                If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                    Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetPath = Project.Parameter("ProNetPath").Value
                Else
                    ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                End If
                Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
                Message.Show() 'Added 18May19
            End If
        Else 'Project has been opened using Command Line arguments.
            Project.ReadParameters()
            Project.ReadParentParameters()
            If Project.ParentParameterExists("ProNetName") Then
                Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                ProNetName = Project.Parameter("ProNetName").Value
            Else
                ProNetName = Project.GetParameter("ProNetName")
            End If
            If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                ProNetPath = Project.Parameter("ProNetPath").Value
            Else
                ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
            End If
            Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

            Project.LockProject() 'Lock the project while it is open in this application.
            ProjectSelected = False 'Reset the Project Selected flag.

            'Set up the Message object:
            Message.SettingsLocn = Project.SettingsLocn
            Message.Show() 'Added 18May19
        End If

        'START Initialise the form: ===============================================================

        Me.WebBrowser1.ObjectForScripting = Me
        'IF THE LINE ABOVE PRODUCES AN ERROR ON STARTUP, CHECK THAT THE CODE ON THE FOLLOWING THREE LINES IS INSERTED JUST ABOVE THE Public Class Main STATEMENT.
        'Imports System.Security.Permissions
        '<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
        '<System.Runtime.InteropServices.ComVisibleAttribute(True)>

        'Initialise Input Data Tab ----------------------------------------------------------
        cmbDatabaseType.Items.Add("Access2007To2013")
        cmbDatabaseType.SelectedIndex = 0 'Select the first item

        'Initialise Titles Tab --------------------------------------------------------------
        Dim items As Array
        items = System.Enum.GetNames(GetType(ContentAlignment))
        Dim item As String
        For Each item In items
            cmbAlignment.Items.Add(item)
        Next

        items = System.Enum.GetNames(GetType(DataVisualization.Charting.TextOrientation))
        For Each item In items
            cmbOrientation.Items.Add(item)
        Next

        'Initialise Chart Settings Tab ------------------------------------------------------

        '  Titles Tab
        txtTitlesRecordNo.Text = "0"
        txtNTitlesRecords.Text = "0"
        '  Series Tab
        txtSeriesRecordNo.Text = "0"
        txtNSeriesRecords.Text = "0"
        '  Areas Tab
        txtAreaRecordNo.Text = "0"
        txtNAreaRecords.Text = "0"

        'Set up the Y Values grid:
        DataGridView1.ColumnCount = 1
        DataGridView1.RowCount = 1
        DataGridView1.Columns(0).HeaderText = "Y Value"
        DataGridView1.Columns(0).Width = 120
        'Dim cboFieldSelections As New DataGridViewComboBoxColumn 'Used for selecting Y Value fields in the Chart Settings tab (This is declared in the Variables Decl section obecause it may be modified later.)
        DataGridView1.Columns.Insert(1, cboFieldSelections)
        DataGridView1.Columns(1).HeaderText = "Field"
        DataGridView1.Columns(1).Width = 120
        DataGridView1.AllowUserToResizeColumns = True

        'Set up the Custom Attributes grid:
        DataGridView2.ColumnCount = 3
        DataGridView2.RowCount = 1
        DataGridView2.Columns(0).HeaderText = "Custom Attribute"
        DataGridView2.Columns(0).Width = 120
        DataGridView2.Columns(1).HeaderText = "Value Range"
        DataGridView2.Columns(1).Width = 120
        DataGridView2.Columns(2).HeaderText = "Value"
        DataGridView2.Columns(2).Width = 120
        DataGridView2.AllowUserToResizeColumns = True

        cmbXAxisType.Items.Add("Primary")
        cmbXAxisType.Items.Add("Secondary")

        For Each item In [Enum].GetNames(GetType(DataVisualization.Charting.ChartValueType))
            cmbXAxisValueType.Items.Add(item)
        Next

        For Each item In [Enum].GetNames(GetType(DataVisualization.Charting.ChartValueType))
            cmbYAxisValueType.Items.Add(item)
        Next

        cmbYAxisType.Items.Add("Primary")
        cmbYAxisType.Items.Add("Secondary")

        cmbXAxisTitleAlignment.Items.Add("Center")
        cmbXAxisTitleAlignment.Items.Add("Far")
        cmbXAxisTitleAlignment.Items.Add("Near")

        cmbX2AxisTitleAlignment.Items.Add("Center")
        cmbX2AxisTitleAlignment.Items.Add("Far")
        cmbX2AxisTitleAlignment.Items.Add("Near")

        cmbYAxisTitleAlignment.Items.Add("Center")
        cmbYAxisTitleAlignment.Items.Add("Far")
        cmbYAxisTitleAlignment.Items.Add("Near")

        cmbY2AxisTitleAlignment.Items.Add("Center")
        cmbY2AxisTitleAlignment.Items.Add("Far")
        cmbY2AxisTitleAlignment.Items.Add("Near")

        SetupStockChartSeriesTab()

        bgwSendMessage.WorkerReportsProgress = True
        bgwSendMessage.WorkerSupportsCancellation = True

        InitialiseForm() 'Initialise the form for a new project.

        'END   Initialise the form: ---------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        OpenStartPage()
        Message.ShowXMessages = ShowXMessages
        Message.ShowSysMessages = ShowSysMessages
        RestoreProjectSettings() 'Restore the Project settings

        ShowProjectInfo() 'Show the project information.

        If chkAutoDraw.Checked Then DrawStockChart()

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

        If StartupConnectionName = "" Then
            If Project.ConnectOnOpen Then
                ConnectToComNet() 'The Project is set to connect when it is opened.
            ElseIf ApplicationInfo.ConnectOnStartup Then
                ConnectToComNet() 'The Application is set to connect when it is started.
            Else
                'Don't connect to ComNet.
            End If

        Else
            'Connect to ComNet using the connection name StartupConnectionName.
            ConnectToComNet(StartupConnectionName)
        End If

        'Get the Application Version Information:
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            'Application is network deployed.
            ApplicationInfo.Version.Number = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            ApplicationInfo.Version.Major = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Major
            ApplicationInfo.Version.Minor = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Minor
            ApplicationInfo.Version.Build = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Build
            ApplicationInfo.Version.Revision = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.Revision
            ApplicationInfo.Version.Source = "Publish"
            Message.Add("Application version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString & vbCrLf)
        Else
            'Application is not network deployed.
            ApplicationInfo.Version.Number = My.Application.Info.Version.ToString
            ApplicationInfo.Version.Major = My.Application.Info.Version.Major
            ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
            ApplicationInfo.Version.Build = My.Application.Info.Version.Build
            ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision
            ApplicationInfo.Version.Source = "Assembly"
            Message.Add("Application version: " & My.Application.Info.Version.ToString & vbCrLf)
        End If

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.
        'OpenStartPage()

        ChartInfo.DataLocation = Project.DataLocn
        If Project.DataLocn.FileExists("LastChart.StockChart") Then
            Try
                ChartInfo.LoadFile("LastChart.StockChart", Chart1)
                UpdateInputDataTabSettings()
                UpdateTitlesTabSettings()
                UpdateSeriesTabSettings()
                UpdateAreasTabSettings()
            Catch ex As Exception
                Message.AddWarning("Error loading LastChart.StockChart: " & ex.Message & vbCrLf & vbCrLf)
            End Try
        End If

    End Sub

    Private Sub ShowProjectInfo()
        'Show the project information:

        txtParentProject.Text = Project.ParentProjectName
        txtProNetName.Text = Project.GetParameter("ProNetName")
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select
        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")

        txtProjectPath.Text = Project.Path

        If Project.ConnectOnOpen Then
            chkConnect.Checked = True
        Else
            chkConnect.Checked = False
        End If

        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsPath.Text = Project.SettingsLocn.Path

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataPath.Text = Project.DataLocn.Path

        Select Case Project.SystemLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSystemLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSystemLocationType.Text = "Archive"
        End Select
        txtSystemPath.Text = Project.SystemLocn.Path

        'txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                        Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                        Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                        Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                        Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                        Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                        Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        DisconnectFromComNet() 'Disconnect from the Communication Network.

        'Update the StockChart settings before saving (The settings may have been changed on the form.)
        'UpdateStockChartSettings()

        SaveProjectSettings() 'Save project settings.

        ApplicationInfo.WriteFile() 'Update the Application Information file.

        Project.SaveLastProjectInfo() 'Save information about the last project used.

        Project.SaveParameters()
        ChartInfo.SaveFile("LastChart.StockChart", Chart1) 'Save the last line chart settings.

        'Project.SaveProjectInfoFile() 'Update the Project Information file. This is not required unless there is a change made to the project.

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.
        ApplicationInfo.UnlockApplication()

        Application.Exit()

    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub


#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.ShowXMessages = ShowXMessages
        Message.MessageForm.BringToFront()
    End Sub


    Private Sub btnWebPages_Click(sender As Object, e As EventArgs) Handles btnWebPages.Click
        'Open the Web Pages form.

        If IsNothing(WebPageList) Then
            WebPageList = New frmWebPageList
            WebPageList.Show()
        Else
            WebPageList.Show()
            WebPageList.BringToFront()
        End If
    End Sub

    Private Sub WebPageList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebPageList.FormClosed
        WebPageList = Nothing
    End Sub

    Public Function OpenNewWebPage() As Integer
        'Open a new HTML Web View window, or reuse an existing one if avaiable.
        'The new forms index number in WebViewFormList is returned.

        NewWebPage = New frmWebPage
        If WebPageFormList.Count = 0 Then
            WebPageFormList.Add(NewWebPage)
            WebPageFormList(0).FormNo = 0
            WebPageFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebPageFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebPageFormList(I)) Then
                    WebPageFormList(I) = NewWebPage
                    WebPageFormList(I).FormNo = I
                    WebPageFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebPageFormList.Add(NewWebPage)
                FormNo = WebPageFormList.Count - 1
                WebPageFormList(FormNo).FormNo = FormNo
                WebPageFormList(FormNo).Show
                Return FormNo 'The new WebPage is at position FormNo in WebPageFormList()
            End If
        End If
    End Function

    Public Sub WebPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If WebPageFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(WebPageFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebPageFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Function OpenNewHtmlDisplayPage() As Integer
        'Open a new HTML display window, or reuse an existing one if avaiable.
        'The new forms index number in HtmlDisplayFormList is returned.

        NewHtmlDisplay = New frmHtmlDisplay
        If HtmlDisplayFormList.Count = 0 Then
            HtmlDisplayFormList.Add(NewHtmlDisplay)
            HtmlDisplayFormList(0).FormNo = 0
            HtmlDisplayFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in HtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in HtmlDisplayFormList. They can be re-used.
                If IsNothing(HtmlDisplayFormList(I)) Then
                    HtmlDisplayFormList(I) = NewHtmlDisplay
                    HtmlDisplayFormList(I).FormNo = I
                    HtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in HtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to HtmlDisplayFormList
                Dim FormNo As Integer
                HtmlDisplayFormList.Add(NewHtmlDisplay)
                FormNo = HtmlDisplayFormList.Count - 1
                HtmlDisplayFormList(FormNo).FormNo = FormNo
                HtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new HtmlDisplay is at position FormNo in HtmlDisplayFormList()
            End If
        End If
    End Function

    Private Sub btnDesignQuery_Click(sender As Object, e As EventArgs) Handles btnDesignQuery.Click
        'Open the Design Query form:
        If IsNothing(DesignQuery) Then
            DesignQuery = New frmDesignQuery
            DesignQuery.Show()
            DesignQuery.DatabasePath = ChartInfo.InputDatabasePath
        Else
            DesignQuery.Show()
            DesignQuery.DatabasePath = ChartInfo.InputDatabasePath
        End If
    End Sub

    Private Sub DesignQuery_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DesignQuery.FormClosed
        DesignQuery = Nothing
    End Sub

    Private Sub btnZoomChart_Click(sender As Object, e As EventArgs) Handles btnZoomChart.Click
        'Open the Zoom Chart form.
        If IsNothing(ZoomChart) Then
            ZoomChart = New frmZoomChart
            ZoomChart.Show()
            ZoomChart.Chart = Chart1
            ZoomChart.SelectAxis(txtAreaName.Text, "X Axis")
        Else
            ZoomChart.Show()
            ZoomChart.SelectAxis(txtAreaName.Text, "X Axis")
        End If
    End Sub

    Private Sub ZoomChart_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ZoomChart.FormClosed
        ZoomChart = Nothing
    End Sub

    Private Sub Chart1_AxisViewChanged(sender As Object, e As ViewEventArgs) Handles Chart1.AxisViewChanged
        If IsNothing(ZoomChart) Then
            ' Message.Add("ZoomChart is Nothing" & vbCrLf)
        Else
            ZoomChart.UpdateSettings() 'Update the Zoom settings. These may have changed if the chart was scrolled.
            'Message.Add("ZoomChart.UpdateSettings()" & vbCrLf)
        End If
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub ApplicationInfo_UpdateExePath() Handles ApplicationInfo.UpdateExePath
        'Update the Executable Path.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath
    End Sub

    Private Sub ApplicationInfo_RestoreDefaults() Handles ApplicationInfo.RestoreDefaults
        'Restore the default application settings.
        DefaultAppProperties()
    End Sub

    Public Sub UpdateWebPage(ByVal FileName As String)
        'Update the web page in WebPageFormList if the Web file name is FileName.

        Dim NPages As Integer = WebPageFormList.Count
        Dim I As Integer

        Try
            For I = 0 To NPages - 1
                If IsNothing(WebPageFormList(I)) Then
                    'Web page has been deleted!
                Else
                    If WebPageFormList(I).FileName = FileName Then
                        WebPageFormList(I).OpenDocument
                    End If
                End If
            Next
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub


#Region " Start Page Code" '=========================================================================================================================================

    Public Sub OpenStartPage()
        'Open the workflow page:

        If Project.DataFileExists(WorkflowFileName) Then
            'Note: WorkflowFileName should have been restored when the application started.
            DisplayWorkflow()
        ElseIf Project.DataFileExists("StartPage.html") Then
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        Else
            CreateStartPage()
            WorkflowFileName = "StartPage.html"
            DisplayWorkflow()
        End If

        ''Open the StartPage.html file and display in the Workflow tab.
        'If Project.DataFileExists("StartPage.html") Then
        '    WorkflowFileName = "StartPage.html"
        '    DisplayWorkflow()
        'Else
        '    CreateStartPage()
        '    WorkflowFileName = "StartPage.html"
        '    DisplayWorkflow()
        'End If
    End Sub

    Public Sub DisplayWorkflow()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(WorkflowFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(WorkflowFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            WebBrowser1.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & WorkflowFileName & vbCrLf)
        End If
    End Sub

    Private Sub CreateStartPage()
        'Create a new default StartPage.html file.

        Dim htmData As New IO.MemoryStream
        Dim sw As New IO.StreamWriter(htmData)
        sw.Write(AppInfoHtmlString("Application Information")) 'Create a web page providing information about the application.
        sw.Flush()
        Project.SaveData("StartPage.html", htmData)
    End Sub

    Public Function AppInfoHtmlString(ByVal DocumentTitle As String) As String
        'Create an Application Information Web Page.

        'This function should be edited to provide a brief description of the Application.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Application information."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Stock Chart" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>" & vbCrLf) 'Add a horizontal divider line.
        sb.Append("<p>The Stock Chart application plots daily stock price data.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<hr>" & vbCrLf & vbCrLf) 'Add a horizontal divider line.

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultJavaScriptString() As String
        'Generate the default JavaScript section of an Andorville(TM) Workflow Web Page.

        Dim sb As New System.Text.StringBuilder

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        Return sb.ToString

    End Function


    Public Function DefaultHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<!-- Andorville(TM) Workflow File -->" & vbCrLf)
        sb.Append("<!-- Application Name:    " & ApplicationInfo.Name & " -->" & vbCrLf)
        sb.Append("<!-- Application Version: " & My.Application.Info.Version.ToString & " -->" & vbCrLf)
        sb.Append("<!-- Creation Date:          " & Format(Now, "dd MMMM yyyy") & " -->" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Workflow description."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h1>" & DocumentTitle & "</h1>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

#End Region 'Start Page Code ------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser1" '==================================
    'These methods are used to display HTML pages in the Document tab.
    'The same methods can be found in the WebView form, which displays web pages on seprate forms.


    'Display Messages ==============================================================================================

    Public Sub AddMessage(ByVal Msg As String)
        'Add a normal text message to the Message window.
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        'Add a warning text message to the Message window.
        Message.AddWarning(Msg)
    End Sub

    Public Sub AddTextTypeMessage(ByVal Msg As String, ByVal TextType As String)
        'Add a message with the specified Text Type to the Message window.
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddXmlMessage(ByVal XmlText As String)
        'Add an Xml message to the Message window.
        Message.AddXml(XmlText)
    End Sub

    'END Display Messages ------------------------------------------------------------------------------------------


    'Run an XSequence ==============================================================================================

    Public Sub RunClipboardXSeq()
        'Run the XSequence instructions in the clipboard.

        Dim XDocSeq As System.Xml.Linq.XDocument
        Try
            XDocSeq = XDocument.Parse(My.Computer.Clipboard.GetText)
        Catch ex As Exception
            Message.AddWarning("Error reading Clipboard data. " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        If IsNothing(XDocSeq) Then
            Message.Add("No XSequence instructions were found in the clipboard.")
        Else
            Dim XmlSeq As New System.Xml.XmlDocument
            Try
                XmlSeq.LoadXml(XDocSeq.ToString) 'Convert XDocSeq to an XmlDocument to process with XSeq.
                'Run the sequence:
                XSeq.RunXSequence(XmlSeq, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XSeq_Instruction(Data As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn
            Case "Settings:Form:Name"
                FormName = Data

            Case "Settings:Form:Item:Name"
                ItemName = Data

            Case "Settings:Form:Item:Value"
                RestoreSetting(FormName, ItemName, Data)

            Case "Settings:Form:SelectId"
                SelectId = Data

            Case "Settings:Form:OptionText"
                RestoreOption(SelectId, Data)


            Case "Settings"

            Case "EndOfSequence"
                'Main.Message.Add("End of processing sequence" & Data & vbCrLf)

            Case Else
                'Main.Message.AddWarning("Unknown location: " & Locn & "  Data: " &  Data & vbCrLf)
                Message.AddWarning("Unknown location: " & Locn & "  Data: " & Data & vbCrLf)

        End Select
    End Sub

    'END Run an XSequence ------------------------------------------------------------------------------------------


    'Run an XMessage ===============================================================================================

    Public Sub RunXMessage(ByVal XMsg As String)
        'Run the XMessage by sending it to InstrReceived.
        InstrReceived = XMsg
    End Sub

    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageExt(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetname.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageWait(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        'Wait for the connection to be made.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            Try
                'Application.DoEvents() 'TRY THE METHOD WITHOUT THE DOEVENTS
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim StartTime As Date = Now
                    Dim Duration As TimeSpan
                    'Wait up to 16 seconds for the connection ConnName to be established
                    While client.ConnectionExists(ProNetName, ConnName) = False 'Wait until the required connection is made.
                        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                        Duration = Now - StartTime
                        If Duration.Seconds > 16 Then Exit While
                    End While

                    If client.ConnectionExists(ProNetName, ConnName) = False Then
                        Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                    Else
                        If bgwSendMessage.IsBusy Then
                            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            Dim SendMessageParams As New clsSendMessageParams
                            SendMessageParams.ProjectNetworkName = ProNetName
                            SendMessageParams.ConnectionName = ConnName
                            SendMessageParams.Message = XMsg
                            bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            If ShowXMessages Then
                                Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Message.XAddXml(XMsg)
                                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub SendXMessageExtWait(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetName.
        'Wait for the connection to be made.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Dim StartTime As Date = Now
                Dim Duration As TimeSpan
                'Wait up to 16 seconds for the connection ConnName to be established
                While client.ConnectionExists(ProNetName, ConnName) = False
                    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    Duration = Now - StartTime
                    If Duration.Seconds > 16 Then Exit While
                End While

                If client.ConnectionExists(ProNetName, ConnName) = False Then
                    Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                Else
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        Dim SendMessageParams As New clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = XMsg
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                        If ShowXMessages Then
                            Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                            Message.XAddXml(XMsg)
                            Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub XMsgInstruction(ByVal Info As String, ByVal Locn As String)
        'Send the XMessage Instruction to the JavaScript function XMsgInstruction for processing.
        Me.WebBrowser1.Document.InvokeScript("XMsgInstruction", New String() {Info, Locn})
    End Sub

    'END Run an XMessage -------------------------------------------------------------------------------------------


    'Get Information ===============================================================================================

    Public Function GetFormNo() As String
        'Return the Form Number of the current instance of the WebPage form.
        'Return FormNo.ToString
        Return "-1" 'The Main Form is not a Web Page form.
    End Function

    Public Function GetParentFormNo() As String
        'Return the Form Number of the Parent Form (that called this form).
        'Return ParentWebPageFormNo.ToString
        Return "-1" 'The Main Form does not have a Parent Web Page.
    End Function

    Public Function GetConnectionName() As String
        'Return the Connection Name of the Project.
        Return ConnectionName
    End Function

    Public Function GetProNetName() As String
        'Return the Project Network Name of the Project.
        Return ProNetName
    End Function

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
    End Sub

    Public Sub ProjectNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Project Network:
        RestoreSetting(FormName, ItemName, Project.Parameter("ProNetName").Value)
    End Sub

    'END Get Information -------------------------------------------------------------------------------------------


    'Open a Web Page ===============================================================================================

    Public Sub OpenWebPage(ByVal FileName As String)
        'Open the web page with the specified File Name.

        If FileName = "" Then

        Else
            'First check if the HTML file is already open:
            Dim FileFound As Boolean = False
            If WebPageFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To WebPageFormList.Count - 1
                    If WebPageFormList(I) Is Nothing Then

                    Else
                        If WebPageFormList(I).FileName = FileName Then
                            FileFound = True
                            WebPageFormList(I).BringToFront
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = OpenNewWebPage()
                WebPageFormList(FormNo).FileName = FileName
                WebPageFormList(FormNo).OpenDocument
                WebPageFormList(FormNo).BringToFront
            End If
        End If
    End Sub

    'END Open a Web Page -------------------------------------------------------------------------------------------


    'Open and Close Projects =======================================================================================

    Public Sub OpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        End If
    End Sub

    Public Sub CheckOpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Relative Path is open.
        'Open it if it is not already open.
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        End If
    End Sub

    Public Sub OpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Path (relative to the Project Network Path) using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CheckOpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Path (relative to the Project Network Path) is open.
        'Open it if it is not already open.
        'Open the Project at the specified Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                'client.StartProjectAtPath(ProjectPath, ConnectionName)
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                'client.StartProjectAtPath(ProjectPath, ConnectionName)
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CloseProjectAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
        'Close the Project at the specified connection.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to close the application at the connection.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                'NOTE: No reply expected. No need to provide the following client information(?)
                'Dim clientConnName As New XElement("ClientConnectionName", Me.ConnectionName)
                'xmessage.Add(clientConnName)

                Dim command As New XElement("Command", "Close")
                xmessage.Add(command)
                doc.Add(xmessage)

                'Show the message sent to AppNet:
                Message.XAddText("Message sent to: [" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage(ProNetName, ConnectionName, doc.ToString)
            End If
        End If
    End Sub

    'END Open and Close Projects -----------------------------------------------------------------------------------


    'System Methods ================================================================================================

    Public Sub SaveHtmlSettings(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:
        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Project.SaveXmlData(FileName, XDocSettings)
    End Sub

    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = WorkflowFileName & "Settings"
        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)
                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RestoreSetting(ByVal FormName As String, ByVal ItemName As String, ByVal ItemValue As String)
        'Restore the setting value with the specified Form Name and Item Name.
        Me.WebBrowser1.Document.InvokeScript("RestoreSetting", New String() {FormName, ItemName, ItemValue})
    End Sub

    Public Sub RestoreOption(ByVal SelectId As String, ByVal OptionText As String)
        'Restore the Option text in the Select control with the Id SelectId.
        Me.WebBrowser1.Document.InvokeScript("RestoreOption", New String() {SelectId, OptionText})
    End Sub

    Private Sub SaveWebPageSettings()
        'Call the SaveSettings JavaScript function:
        Try
            Me.WebBrowser1.Document.InvokeScript("SaveSettings")
        Catch ex As Exception
            Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try
    End Sub

    'END System Methods --------------------------------------------------------------------------------------------


    'Legacy Code (These methods should no longer be used) ==========================================================

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    Public Sub ShowMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    'END Legacy Code -----------------------------------------------------------------------------------------------


#End Region 'Methods Called by JavaScript -------------------------------------------------------------------------------------------------------------------------------


#Region " Project Events Code"

    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg & vbCrLf)
    End Sub

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
        Message.AddWarning(Msg & vbCrLf)
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.
        CloseProject()
        'SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        'SaveProjectSettings() 'Update this subroutine if project settings need to be saved.
        'Project.Usage.SaveUsageInfo() 'Save the current project usage information.
        'Project.UnlockProject() 'Unlock the current project before it Is closed.
        'If ConnectedToComNet Then DisconnectFromComNet()
    End Sub

    Private Sub CloseProject()
        'Close the Project:
        SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        SaveProjectSettings() 'Update this subroutine if project settings need to be saved.
        Project.Usage.SaveUsageInfo() 'Save the current project usage information.
        Project.UnlockProject() 'Unlock the current project before it Is closed.
        If ConnectedToComNet Then DisconnectFromComNet() 'ADDED 9Apr20
    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.
        OpenProject()
        'RestoreFormSettings()
        'Project.ReadProjectInfoFile()

        'Project.ReadParameters()
        'Project.ReadParentParameters()
        'If Project.ParentParameterExists("ProNetName") Then
        '    Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
        '    ProNetName = Project.Parameter("ProNetName").Value
        'Else
        '    ProNetName = Project.GetParameter("ProNetName")
        'End If
        'If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
        '    Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
        '    ProNetPath = Project.Parameter("ProNetPath").Value
        'Else
        '    ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
        'End If
        'Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

        'Project.LockProject() 'Lock the project while it is open in this application.

        'Project.Usage.StartTime = Now

        'ApplicationInfo.SettingsLocn = Project.SettingsLocn
        'Message.SettingsLocn = Project.SettingsLocn
        'Message.Show() 'Added 18May19

        ''Restore the new project settings:
        'RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        'ShowProjectInfo()

        '''Show the project information:
        ''txtProjectName.Text = Project.Name
        ''txtProjectDescription.Text = Project.Description
        ''Select Case Project.Type
        ''    Case ADVL_Utilities_Library_1.Project.Types.Directory
        ''        txtProjectType.Text = "Directory"
        ''    Case ADVL_Utilities_Library_1.Project.Types.Archive
        ''        txtProjectType.Text = "Archive"
        ''    Case ADVL_Utilities_Library_1.Project.Types.Hybrid
        ''        txtProjectType.Text = "Hybrid"
        ''    Case ADVL_Utilities_Library_1.Project.Types.None
        ''        txtProjectType.Text = "None"
        ''End Select

        ''txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        ''txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        ''Select Case Project.SettingsLocn.Type
        ''    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ''        txtSettingsLocationType.Text = "Directory"
        ''    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        ''        txtSettingsLocationType.Text = "Archive"
        ''End Select
        ''txtSettingsPath.Text = Project.SettingsLocn.Path
        ''Select Case Project.DataLocn.Type
        ''    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ''        txtDataLocationType.Text = "Directory"
        ''    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        ''        txtDataLocationType.Text = "Archive"
        ''End Select
        ''txtDataPath.Text = Project.DataLocn.Path

        'If Project.ConnectOnOpen Then
        '    ConnectToComNet() 'The Project is set to connect when it is opened.
        'ElseIf ApplicationInfo.ConnectOnStartup Then
        '    ConnectToComNet() 'The Application is set to connect when it is started.
        'Else
        '    'Don't connect to ComNet.
        'End If

    End Sub

    Private Sub OpenProject()
        'Open the Project:
        RestoreFormSettings()
        Project.ReadProjectInfoFile()

        Project.ReadParameters()
        Project.ReadParentParameters()
        If Project.ParentParameterExists("ProNetName") Then
            Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
            ProNetName = Project.Parameter("ProNetName").Value
        Else
            ProNetName = Project.GetParameter("ProNetName")
        End If
        If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
            Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
            ProNetPath = Project.Parameter("ProNetPath").Value
        Else
            ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
        End If
        Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

        Project.LockProject() 'Lock the project while it is open in this application.

        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show() 'Added 18May19

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        ShowProjectInfo()

        If Project.ConnectOnOpen Then
            ConnectToComNet() 'The Project is set to connect when it is opened.
        ElseIf ApplicationInfo.ConnectOnStartup Then
            ConnectToComNet() 'The Application is set to connect when it is started.
        Else
            'Don't connect to ComNet.
        End If
    End Sub

#End Region 'Project Events Code

#Region " Online/Offline Code" '=========================================================================================================================================

    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        'Connect to or disconnect from the Message System (ComNet).
        If ConnectedToComNet = False Then
            ConnectToComNet()
        Else
            DisconnectFromComNet()
        End If
    End Sub

    Private Sub ConnectToComNet()
        'Connect to the Message System. (ComNet)

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If


        'UPDATE 14 Feb 2021 - If the VS2019 version of the ADVL Network is running it may not detected by ComNetRunning()!
        'Check if the Message Service is running by trying to open a connection:
        Try
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)
            ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
            ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)
            If ConnectionName <> "" Then
                Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                btnOnline.Text = "Online"
                btnOnline.ForeColor = Color.ForestGreen
                ConnectedToComNet = True
                SendApplicationInfo()
                SendProjectInfo()
                client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                bgwComCheck.WorkerReportsProgress = True
                bgwComCheck.WorkerSupportsCancellation = True
                If bgwComCheck.IsBusy Then
                    'The ComCheck thread is already running.
                Else
                    bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                End If
                Exit Sub 'Connection made OK
            Else
                'Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                Message.Add("The Andorville™ Network was not found. Attempting to start it." & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            End If
        Catch ex As System.TimeoutException
            Message.Add("Message Service Check Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        Catch ex As Exception
            Message.Add("Error message: " & ex.Message & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        End Try

        If ComNetRunning() Then
            'The Application.Lock file has been found at AdvlNetworkAppPath
            'The Message Service is Running.
        Else 'The Message Service is NOT running'
            'Start the Message Service:
            If AdvlNetworkAppPath = "" Then
                Message.AddWarning("Andorville™ Network application path is unknown." & vbCrLf)
            Else
                If System.IO.File.Exists(AdvlNetworkExePath) Then 'OK to start the Message Service application:
                    Shell(Chr(34) & AdvlNetworkExePath & Chr(34), AppWinStyle.NormalFocus) 'Start Message Service application with no argument
                Else
                    'Incorrect Message Service Executable path.
                    Message.AddWarning("Andorville™ Network exe file not found. Service not started." & vbCrLf)
                End If
            End If
        End If

        'Try to fix a faulted client state:
        If client.State = ServiceModel.CommunicationState.Faulted Then
            client = Nothing
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds 8 seconds is too short for a slow computer!)

                ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
                ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                If ConnectionName <> "" Then
                    Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToComNet = True
                    SendApplicationInfo()
                    SendProjectInfo()
                    client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                    bgwComCheck.WorkerReportsProgress = True
                    bgwComCheck.WorkerSupportsCancellation = True
                    If bgwComCheck.IsBusy Then
                        'The ComCheck thread is already running.
                    Else
                        bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                    End If

                Else
                    Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If
    End Sub

    Private Sub ConnectToComNet(ByVal ConnName As String)
        'Connect to the Communication Network with the connection name ConnName.


        'UPDATE 14 Feb 2021 - If the VS2019 version of the ADVL Network is running it may not be detected by ComNetRunning()!
        'Check if the Message Service is running by trying to open a connection:

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        Try
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)
            ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
            ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)
            If ConnectionName <> "" Then
                Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                btnOnline.Text = "Online"
                btnOnline.ForeColor = Color.ForestGreen
                ConnectedToComNet = True
                SendApplicationInfo()
                SendProjectInfo()
                client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                bgwComCheck.WorkerReportsProgress = True
                bgwComCheck.WorkerSupportsCancellation = True
                If bgwComCheck.IsBusy Then
                    'The ComCheck thread is already running.
                Else
                    bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                End If
                Exit Sub 'Connection made OK
            Else
                'Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                Message.Add("The Andorville™ Network was not found. Attempting to start it." & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            End If
        Catch ex As System.TimeoutException
            Message.Add("Message Service Check Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        Catch ex As Exception
            Message.Add("Error message: " & ex.Message & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        End Try


        If ConnectedToComNet = False Then
            'Dim Result As Boolean

            If IsNothing(client) Then
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            'Try to fix a faulted client state:
            If client.State = ServiceModel.CommunicationState.Faulted Then
                client = Nothing
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.AddWarning("client state is faulted. Connection not made!" & vbCrLf)
            Else
                Try
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)
                    ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
                    ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                    If ConnectionName <> "" Then
                        Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                        btnOnline.Text = "Online"
                        btnOnline.ForeColor = Color.ForestGreen
                        ConnectedToComNet = True
                        SendApplicationInfo()
                        SendProjectInfo()
                        client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                        bgwComCheck.WorkerReportsProgress = True
                        bgwComCheck.WorkerSupportsCancellation = True
                        If bgwComCheck.IsBusy Then
                            'The ComCheck thread is already running.
                        Else
                            bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                        End If

                    Else
                        Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    End If
                Catch ex As System.TimeoutException
                    Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
                Catch ex As Exception
                    Message.Add("Error message: " & ex.Message & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End Try
            End If
        Else
            Message.AddWarning("Already connected to the Andorville™ Network (Message Service)." & vbCrLf)
        End If
    End Sub

    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network.

        If ConnectedToComNet = True Then
            If IsNothing(client) Then
                'Message.Add("Already disconnected from the Communication Network." & vbCrLf)
                Message.Add("Already disconnected from the Andorville™ Network (Message Service)." & vbCrLf)
                btnOnline.Text = "Offline"
                btnOnline.ForeColor = Color.Red
                ConnectedToComNet = False
                ConnectionName = ""
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted." & vbCrLf)
                    ConnectionName = ""
                Else
                    Try
                        'client.Disconnect(AppNetName, ConnectionName)
                        client.Disconnect(ProNetName, ConnectionName)
                        btnOnline.Text = "Offline"
                        btnOnline.ForeColor = Color.Red
                        ConnectedToComNet = False
                        ConnectionName = ""
                        'Message.Add("Disconnected from the Communication Network." & vbCrLf)
                        Message.Add("Disconnected from the Andorville™ Network (Message Service)." & vbCrLf)

                        If bgwComCheck.IsBusy Then
                            bgwComCheck.CancelAsync()
                        End If
                    Catch ex As Exception
                        'Message.AddWarning("Error disconnecting from Communication Network: " & ex.Message & vbCrLf)
                        Message.AddWarning("Error disconnecting from Andorville™ Network (Message Service): " & ex.Message & vbCrLf)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub SendApplicationInfo()
        'Send the application information to the Administrator connections.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to send application information.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                Dim applicationInfo As New XElement("ApplicationInfo")
                Dim name As New XElement("Name", Me.ApplicationInfo.Name)
                applicationInfo.Add(name)

                Dim text As New XElement("Text", "Stock Chart")
                applicationInfo.Add(text)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)

                'Show the message sent to ComNet:
                Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage("", "MessageService", doc.ToString)

            End If
        End If

    End Sub

    Private Sub SendProjectInfo()
        'Send the project information to the Network application.

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Public Sub SendProjectInfo(ByVal ProjectPath As String)
        'Send the project information to the Network application.
        'This version of SendProjectInfo uses the ProjectPath argument.

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", ProjectPath)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Private Function ComNetRunning() As Boolean
        'Return True if ComNet (Message Service) is running.
        'If System.IO.File.Exists(MsgServiceAppPath & "\Application.Lock") Then
        '    Return True
        'Else
        '    Return False
        'End If
        If AdvlNetworkAppPath = "" Then
            Message.Add("Andorville™ Network application path is not known." & vbCrLf)
            Message.Add("Run the MAndorville™ Network before connecting to update the path." & vbCrLf)
            Return False
        Else
            If System.IO.File.Exists(AdvlNetworkAppPath & "\Application.Lock") Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

#End Region 'Online/Offline code ----------------------------------------------------------------------------------------------------------------------------------------

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        'Update the current duration:
        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                                   Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                                   Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                                   Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:
        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
        '                          Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                           Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                           Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                           Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Add the current project to the Message Service list.

        If Project.ParentProjectName <> "" Then
            Message.AddWarning("This project has a parent: " & Project.ParentProjectName & vbCrLf)
            Message.AddWarning("Child projects can not be added to the list." & vbCrLf)
            Exit Sub
        End If

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to AppNet:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub



#Region " Process XMessages" '===========================================================================================================================================

    Private Sub XMsg_Instruction(Data As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville™ applications.
        '
        'An XSequence file is an AL-H7™ Information Sequence stored in an XML format.
        'AL-H7™ is the name of a programming system that uses sequences of data and location value pairs to store information or processing steps.
        'Any program, mathematical expression or data set can be expressed as an Information Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville™ applications for examples.

        If IsDBNull(Data) Then
            Data = ""
        End If

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Data, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn

            'Case "ClientAppNetName"
            '    ClientAppNetName = Data 'The name of the Client Application Network requesting service. ADDED 2Feb19.
                Case "ClientProNetName"
                    ClientProNetName = Data 'The name of the Client Application Network requesting service. 

                Case "ClientName"
                    ClientAppName = Data 'The name of the Client requesting service.

                Case "ClientConnectionName"
                    ClientConnName = Data 'The name of the client requesting service.

                Case "ClientLocn" 'The Location within the Client requesting service.
                    Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                    xlocns.Add(New XElement(Data)) 'Start the new location instructions

                'Case "OnCompletion" 'Specify the last instruction to be returned on completion of the XMessage processing.
                '    CompletionInstruction = Data

                       'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data

                Case "Main"
                 'Blank message - do nothing.

                'Case "Main:OnCompletion"
                '    Select Case "Stop"
                '        'Stop on completion of the instruction sequence.
                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                    End Select

                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "Command"
                    Select Case Data
                        Case "ConnectToComNet" 'Startup Command
                            If ConnectedToComNet = False Then
                                ConnectToComNet()
                            End If
                        Case "GetStockChartSettings"
                        'GetStockChartSettingsForClient() 'NOW USING: GetChartSettings

                        Case "AppComCheck"
                            'Add the Appplication Communication info to the reply message:
                            Dim clientProNetName As New XElement("ClientProNetName", ProNetName) 'The Project Network Name
                            xlocns(xlocns.Count - 1).Add(clientProNetName)
                            Dim clientName As New XElement("ClientName", "ADVL_Stock_Chart_1") 'The name of this application.
                            xlocns(xlocns.Count - 1).Add(clientName)
                            Dim clientConnectionName As New XElement("ClientConnectionName", ConnectionName)
                            xlocns(xlocns.Count - 1).Add(clientConnectionName)
                            '<Status>OK</Status> will be automatically appended to the XMessage before it is sent.

                    End Select

                Case "GetChartSettings"
                    'This is the new code used to send the current chart settings to the client.
                    '  The chart settings will be returned to the client in an XDataMsg file (instead of an XMsg file).
                    '  Data contains the Client Location to include in the XDataMsg file.
                    GetStockChartSettings(Data)

            ''Stock Chart instructions: --------------------------------------------------------------------------------------------------------------------------------
            ''SEE NEW INSTRUCTIONS!!!

            'Case "StockChartSettings:ChartType"
            '    Select Case Data
            '        Case "Stock"
            '            'This application plots this type of chart.
            '        Case Else
            '            Message.AddWarning("Chart type is: " & Data & vbCrLf)
            '            Message.AddWarning("This application plots 'Stock' charts." & vbCrLf)
            '    End Select

            'Case "StockChartSettings:InputData:Type"
            '    Select Case Data
            '        Case "Database"
            '            rbDatabase.Checked = True
            '            'StockChart.InputDataType = "Database"
            '            SelectedChartInfo.InputDataType = "Database"
            '    End Select

            'Case "StockChartSettings:InputData:DatabasePath"
            '    'InputDatabasePath = Data
            '    'StockChart.InputDatabasePath = Data
            '    SelectedChartData.InputDatabasePath = Data

            'Case "StockChartSettings:InputData:DataDescription"
            '    'InputDataDescr = Data
            '    'StockChart.InputDataDescr = Data
            '    SelectedChartInfo.InputDataDescr = Data

            'Case "StockChartSettings:InputData:DatabaseQuery"
            '    'InputQuery = Data
            '    'StockChart.InputQuery = Data
            '    SelectedChartInfo.InputQuery = Data

            'Case "StockChartSettings:ChartProperties:XValuesFieldName"
            '    'StockChart.XValuesFieldName = Data
            '    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).XValuesFieldName = Data

            'Case "StockChartSettings:ChartProperties:SeriesName"
            '    'StockChart.SeriesName = Data
            '    SeriesInfoName = Data
            '    SelectedChartInfo.dictSeriesInfo.Add(SeriesInfoName, New SeriesInfo)
            '    AreaName = "ChartArea1" 'Specify a chart area name - when the code is updated, a sepaarate instruction will do this.

            'Case "StockChartSettings:ChartProperties:YValuesHighFieldName"
            '    'StockChart.YValuesHighFieldName = Data
            '    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesHighFieldName = Data

            'Case "StockChartSettings:ChartProperties:YValuesLowFieldName"
            '    'StockChart.YValuesLowFieldName = Data
            '    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesLowFieldName = Data

            'Case "StockChartSettings:ChartProperties:YValuesOpenFieldName"
            '    'StockChart.YValuesOpenFieldName = Data
            '    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesOpenFieldName = Data

            'Case "StockChartSettings:ChartProperties:YValuesCloseFieldName"
            '    'StockChart.YValuesCloseFieldName = Data
            '    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesCloseFieldName = Data

            'Case "StockChartSettings:ChartTitle:LabelName"
            '    'StockChart.ChartLabel.Name = Data
            '    TitleName = Data
            '    SelectedChart.Titles.Add(TitleName)
            '    ChartFontStyle = FontStyle.Regular 'Reset the font style when a new title is added.
            '    ChartFontSize = 12                 'Reset the font size when a new title is added.

            'Case "StockChartSettings:ChartTitle:Text"
            '    'StockChart.ChartLabel.Text = Data
            '    SelectedChart.Titles(TitleName).Text = Data

            'Case "StockChartSettings:ChartTitle:FontName"
            '    'StockChart.ChartLabel.FontName = Data
            '    ChartFontName = Data
            '    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Case "StockChartSettings:ChartTitle:Color"
            '    'StockChart.ChartLabel.Color = Data
            '    'StockChart.ChartLabel.Color = Color.FromName(Info)
            '    Try
            '        'StockChart.ChartLabel.Color = Color.FromArgb(Info)
            '        SelectedChart.Titles(TitleName).ForeColor = Color.FromArgb(Info)
            '    Catch ex As Exception

            '    End Try

            'Case "StockChartSettings:ChartTitle:Size"
            '    'StockChart.ChartLabel.Size = Data
            '    ChartFontSize = Data
            '    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Case "StockChartSettings:ChartTitle:Bold"
            '    'StockChart.ChartLabel.Bold = Data
            '    ChartFontStyle = ChartFontStyle Or FontStyle.Bold
            '    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Case "StockChartSettings:ChartTitle:Italic"
            '    'StockChart.ChartLabel.Italic = Data
            '    ChartFontStyle = ChartFontStyle Or FontStyle.Italic
            '    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Case "StockChartSettings:ChartTitle:Underline"
            '    'StockChart.ChartLabel.Underline = Data
            '    ChartFontStyle = ChartFontStyle Or FontStyle.Underline
            '    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Case "StockChartSettings:ChartTitle:Strikeout"
            '    'StockChart.ChartLabel.Strikeout = Data
            '    ChartFontStyle = ChartFontStyle Or FontStyle.Strikeout
            '    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Case "StockChartSettings:ChartTitle:Alignment"
            '    'Select Case Data
            '    '    Case "BottomCenter"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.BottomCenter
            '    '    Case "BottomLeft"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.BottomLeft
            '    '    Case "BottomRight"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.BottomRight
            '    '    Case "MiddleCenter"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.MiddleCenter
            '    '    Case "MiddleLeft"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.MiddleLeft
            '    '    Case "MiddleRight"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.MiddleRight
            '    '    Case "TopCenter"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.TopCenter
            '    '    Case "TopLeft"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.TopLeft
            '    '    Case "TopRight"
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.TopRight
            '    '    Case Else
            '    '        Message.AddWarning("Unknown Chart Title Alignment: " &Data& vbCrLf)
            '    '        StockChart.ChartLabel.Alignment = ContentAlignment.TopCenter
            '    'End Select
            '    SelectedChart.Titles(TitleName).Alignment = Data

            'Case "StockChartSettings:XAxis:TitleText"
            '    'StockChart.XAxis.Title.Text = Data
            '    SelectedChart.ChartAreas(AreaName).AxisX.Title = Data


            'Case "StockChartSettings:XAxis:TitleFontName"
            '    'StockChart.XAxis.Title.FontName = Data
            '    ChartFontName = Data
            '    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle)

            'Case "StockChartSettings:XAxis:TitleColor"
            '    'StockChart.XAxis.Title.Color = Data
            '    'NOTE: Color is stored here as a string!

            '    'Try
            '    '    StockChart.XAxis.Title.Color = Color.FromArgb(Info)
            '    'Catch ex As Exception

            '    'End Try

            '    SelectedChart.ChartAreas(AreaName).AxisX.TitleForeColor = Color.FromArgb(Info)

            'Case "StockChartSettings:XAxis:TitleSize"
            '    StockChart.XAxis.Title.Size = Data

            'Case "StockChartSettings:XAxis:TitleBold"
            '    StockChart.XAxis.Title.Bold = Data

            'Case "StockChartSettings:XAxis:TitleItalic"
            '    StockChart.XAxis.Title.Italic = Data

            'Case "StockChartSettings:XAxis:TitleUnderline"
            '    StockChart.XAxis.Title.Underline = Data

            'Case "StockChartSettings:XAxis:TitleStrikeout"
            '    StockChart.XAxis.Title.Strikeout = Data

            'Case "StockChartSettings:XAxis:TitleAlignment"
            '    Select Case Data
            '        Case "Center"
            '            StockChart.XAxis.TitleAlignment = StringAlignment.Center
            '        Case "Far"
            '            StockChart.XAxis.TitleAlignment = StringAlignment.Far
            '        Case "Near"
            '            StockChart.XAxis.TitleAlignment = StringAlignment.Near
            '        Case Else
            '            Message.AddWarning("Unknown Chart X Axis Title Alignment: " &Data& vbCrLf)
            '            StockChart.XAxis.TitleAlignment = StringAlignment.Center
            '    End Select

            'Case "StockChartSettings:XAxis:AutoMinimum"
            '    StockChart.XAxis.AutoMinimum = Data

            'Case "StockChartSettings:XAxis:Minimum"
            '    StockChart.XAxis.Minimum = Data

            'Case "StockChartSettings:XAxis:AutoMaximum"
            '    StockChart.XAxis.AutoMaximum = Data

            'Case "StockChartSettings:XAxis:Maximum"
            '    StockChart.XAxis.Maximum = Data

            'Case "StockChartSettings:XAxis:AutoInterval"
            '    StockChart.XAxis.AutoInterval = Data

            'Case "StockChartSettings:XAxis:MajorGridInterval"
            '    StockChart.XAxis.MajorGridInterval = Data

            'Case "StockChartSettings:XAxis:AutoMajorGridInterval"
            '    StockChart.XAxis.AutoMajorGridInterval = Data

            'Case "StockChartSettings:YAxis:TitleText"
            '    StockChart.YAxis.Title.Text = Data

            'Case "StockChartSettings:YAxis:TitleFontName"
            '    StockChart.YAxis.Title.FontName = Data

            'Case "StockChartSettings:YAxis:TitleColor"
            '    StockChart.YAxis.Title.Color = Data

            'Case "StockChartSettings:YAxis:TitleSize"
            '    StockChart.YAxis.Title.Size = Data

            'Case "StockChartSettings:YAxis:TitleBold"
            '    StockChart.YAxis.Title.Bold = Data

            'Case "StockChartSettings:YAxis:TitleItalic"
            '    StockChart.YAxis.Title.Italic = Data

            'Case "StockChartSettings:YAxis:TitleUnderline"
            '    StockChart.YAxis.Title.Underline = Data

            'Case "StockChartSettings:YAxis:TitleStrikeout"
            '    StockChart.YAxis.Title.Strikeout = Data

            'Case "StockChartSettings:YAxis:TitleAlignment"
            '    'StockChart.YAxis.TitleAlignment = Data
            '    Select Case Data
            '        Case "Center"
            '            StockChart.YAxis.TitleAlignment = StringAlignment.Center
            '        Case "Far"
            '            StockChart.YAxis.TitleAlignment = StringAlignment.Far
            '        Case "Near"
            '            StockChart.YAxis.TitleAlignment = StringAlignment.Near
            '        Case Else
            '            Message.AddWarning("Unknown Chart Y Axis Title Alignment: " &Data& vbCrLf)
            '            StockChart.YAxis.TitleAlignment = StringAlignment.Center
            '    End Select

            'Case "StockChartSettings:YAxis:AutoMinimum"
            '    StockChart.YAxis.AutoMinimum = Data

            'Case "StockChartSettings:YAxis:Minimum"
            '    StockChart.YAxis.Minimum = Data

            'Case "StockChartSettings:YAxis:AutoMaximum"
            '    StockChart.YAxis.AutoMaximum = Data

            'Case "StockChartSettings:YAxis:Maximum"
            '    StockChart.YAxis.Maximum = Data

            'Case "StockChartSettings:YAxis:AutoInterval"
            '    StockChart.YAxis.AutoInterval = Data

            'Case "StockChartSettings:YAxis:MajorGridInterval"
            '    StockChart.YAxis.MajorGridInterval = Data

            'Case "StockChartSettings:YAxis:AutoMajorGridInterval"
            '    StockChart.YAxis.AutoMajorGridInterval = Data

            'Case "StockChartSettings:Command"
            '    Select Case Data
            '        Case "DrawChart"
            '            UpdateStockChartForm()
            '            'DrawChart()
            '            DrawStockChart()
            '        Case "ClearChart"
            '            'ClearChart()
            '            StockChart.Clear()
            '            UpdateStockChartForm()
            '    End Select



            'Stock Chart instructions: ================================================================================================================================

                Case "ClearChart"
                    SelectedChartInfo.Clear(SelectedChart) 'Clear the chart.

            'Input Data:
                Case "InputDataType"
                    Select Case Data
                        Case "Database"
                            'ChartInfo.InputDataType = "Database"
                            'rbDatabase.Checked = True
                            SelectedChartInfo.InputDataType = "Database"
                    End Select
                Case "InputDatabasePath"
                    SelectedChartInfo.InputDatabasePath = Data
                Case "InputQuery"
                    SelectedChartInfo.InputQuery = Data
                Case "InputDataDescr"
                    SelectedChartInfo.InputDataDescr = Data

            'Series Info:
                Case "SeriesInfo:Name"
                    SeriesInfoName = Data
                    SelectedChartInfo.dictSeriesInfo.Add(SeriesInfoName, New SeriesInfo)
                Case "SeriesInfo:XValuesFieldName"
                    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).XValuesFieldName = Data
                Case "SeriesInfo:YValuesHighFieldName"
                    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesHighFieldName = Data
                Case "SeriesInfo:YValuesLowFieldName"
                    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesLowFieldName = Data
                Case "SeriesInfo:YValuesOpenFieldName"
                    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesOpenFieldName = Data
                Case "SeriesInfo:YValuesCloseFieldName"
                    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).YValuesCloseFieldName = Data
                Case "SeriesInfo:ChartArea"
                    SelectedChartInfo.dictSeriesInfo(SeriesInfoName).ChartArea = Data

            'Area Info:
                Case "AreaInfo:Name"
                    AreaInfoName = Data
                    SelectedChartInfo.dictAreaInfo.Add(AreaInfoName, New AreaInfo)
                Case "AreaInfo:AutoXAxisMinimum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoXAxisMinimum = Data
                Case "AreaInfo:AutoXAxisMaximum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoXAxisMaximum = Data
                Case "AreaInfo:AutoXAxisMajorGridInterval"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoXAxisMajorGridInterval = Data
                Case "AreaInfo:AutoX2AxisMinimum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoX2AxisMinimum = Data
                Case "AreaInfo:AutoX2AxisMaximum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoX2AxisMaximum = Data
                Case "AreaInfo:AutoX2AxisMajorGridInterval"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoX2AxisMajorGridInterval = Data
                Case "AreaInfo:AutoYAxisMinimum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoYAxisMinimum = Data
                Case "AreaInfo:AutoYAxisMaximum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoYAxisMaximum = Data
                Case "AreaInfo:AutoYAxisMajorGridInterval"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoYAxisMajorGridInterval = Data
                Case "AreaInfo:AutoY2AxisMinimum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoY2AxisMinimum = Data
                Case "AreaInfo:AutoY2AxisMaximum"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoY2AxisMaximum = Data
                Case "AreaInfo:AutoY2AxisMajorGridInterval"
                    SelectedChartInfo.dictAreaInfo(AreaInfoName).AutoY2AxisMajorGridInterval = Data

            'Title:
                Case "Title:Name"
                    TitleName = Data
                    SelectedChart.Titles.Add(TitleName).Name = TitleName
                    ChartFontStyle = FontStyle.Regular 'Reset the font style when a new title is added.
                    ChartFontSize = 12                 'Reset the font size when a new title is added.
                Case "Title:Text"
                    SelectedChart.Titles(TitleName).Text = Data
                Case "Title:TextOrientation"
                    SelectedChart.Titles(TitleName).TextOrientation = Data
                Case "Title:Alignment"
                    SelectedChart.Titles(TitleName).Alignment = Data
                Case "Title:ForeColor"
                    SelectedChart.Titles(TitleName).ForeColor = Color.FromArgb(Data)
                Case "Title:Font:Name"
                    ChartFontName = Data
                    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "Title:Font:Size"
                    ChartFontSize = Data
                    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "Title:Font:Bold"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Bold
                    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "Title:Font:Italic"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Italic
                    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "Title:Font:Strikeout"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Strikeout
                    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "Title:Font:Underline"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Underline
                    SelectedChart.Titles(TitleName).Font = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.

            'Series:
                Case "Series:Name"
                    SeriesName = Data
                    SelectedChart.Series.Add(SeriesName)
                Case "Series:ChartType"
                    SelectedChart.Series(SeriesName).ChartType = Data
                Case "Series:Legend"
                    SelectedChart.Series(SeriesName).Legend = Data
                Case "Series:EmptyPointValue"
                    SelectedChart.Series(SeriesName).SetCustomProperty("EmptyPointValue", Data)
                Case "Series:LabelStyle"
                    SelectedChart.Series(SeriesName).SetCustomProperty("LabelStyle", Data)
                Case "Series:PixelPointDepth"
                    SelectedChart.Series(SeriesName).SetCustomProperty("PixelPointDepth", Data)
                Case "Series:PixelPointGapDepth"
                    SelectedChart.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", Data)
                Case "Series:ShowMarkerLines"
                    SelectedChart.Series(SeriesName).SetCustomProperty("ShowMarkerLines", Data)
                Case "Series:AxisLabel"
                    SelectedChart.Series(SeriesName).AxisLabel = Data
                Case "Series:XAxisType"
                    SelectedChart.Series(SeriesName).XAxisType = [Enum].Parse(GetType(DataVisualization.Charting.AxisType), Data)
                Case "Series:XValueType"
                    SelectedChart.Series(SeriesName).XValueType = [Enum].Parse(GetType(DataVisualization.Charting.ChartValueType), Data)
                Case "Series:YAxisType"
                    SelectedChart.Series(SeriesName).YAxisType = [Enum].Parse(GetType(DataVisualization.Charting.AxisType), Data)
                Case "Series:YValueType"
                    SelectedChart.Series(SeriesName).YValueType = [Enum].Parse(GetType(DataVisualization.Charting.ChartValueType), Data)
                Case "Series:Marker:BorderColor"
                    SelectedChart.Series(SeriesName).BorderColor = Color.FromArgb(Data)
                Case "Series:Marker:BorderWidth"
                    SelectedChart.Series(SeriesName).MarkerBorderWidth = Data
                Case "Series:Marker:Color"
                    SelectedChart.Series(SeriesName).MarkerColor = Color.FromArgb(Data)
                Case "Series:Marker:Size"
                    SelectedChart.Series(SeriesName).MarkerSize = Data
                Case "Series:Marker:Step"
                    SelectedChart.Series(SeriesName).MarkerStep = Data
                Case "Series:Marker:Style"
                    SelectedChart.Series(SeriesName).MarkerStyle = [Enum].Parse(GetType(DataVisualization.Charting.MarkerStyle), Data)

            'Chart Area:
                Case "ChartArea:Name"
                    AreaName = Data
                    SelectedChart.ChartAreas.Add(AreaName)
                    ChartFontStyle = FontStyle.Regular 'Reset the font style when a new title is added.
                    ChartFontSize = 12                 'Reset the font size when a new title is added.

            'Chart Area AxisX:
                Case "ChartArea:AxisX:Title:Text"
                    SelectedChart.ChartAreas(AreaName).AxisX.Title = Data
                Case "ChartArea:AxisX:Title:Alignment"
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleAlignment = Data
                Case "ChartArea:AxisX:Title:ForeColor"
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleForeColor = Color.FromArgb(Data)
                Case "ChartArea:AxisX:Title:Font:Name"
                    ChartFontName = Data
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX:Title:Font:Size"
                    ChartFontSize = Data
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX:Title:Font:Bold"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Bold
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX:Title:Font:Italic"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Italic
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX:Title:Font:Strikeout"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Strikeout
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX:Title:Font:Underline"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Underline
                    SelectedChart.ChartAreas(AreaName).AxisX.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX:LabelStyleFormat"
                    SelectedChart.ChartAreas(AreaName).AxisX.LabelStyle.Format = Data
                Case "ChartArea:AxisX:Minimum"
                    SelectedChart.ChartAreas(AreaName).AxisX.Minimum = Data
                Case "ChartArea:AxisX:Maximum"
                    SelectedChart.ChartAreas(AreaName).AxisX.Maximum = Data
                Case "ChartArea:AxisX:LineWidth"
                    SelectedChart.ChartAreas(AreaName).AxisX.LineWidth = Data
                Case "ChartArea:AxisX:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisX.Interval = Data
                Case "ChartArea:AxisX:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisX.IntervalOffset = Data
                Case "ChartArea:AxisX:Crossing"
                    SelectedChart.ChartAreas(AreaName).AxisX.Crossing = Data
                Case "ChartArea:AxisX:MajorGrid:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisX.MajorGrid.Interval = Data
                Case "ChartArea:AxisX:MajorGrid:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisX.MajorGrid.IntervalOffset = Data

            'Chart Area AxisX2
                Case "ChartArea:AxisX2:Title:Text"
                    SelectedChart.ChartAreas(AreaName).AxisX2.Title = Data
                Case "ChartArea:AxisX2:Title:Alignment"
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleAlignment = Data
                Case "ChartArea:AxisX2:Title:ForeColor"
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleForeColor = Color.FromArgb(Data)
                Case "ChartArea:AxisX2:Title:Font:Name"
                    ChartFontName = Data
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX2:Title:Font:Size"
                    ChartFontSize = Data
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX2:Title:Font:Bold"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Bold
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX2:Title:Font:Italic"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Italic
                    Chart1.ChartAreas(AreaName).AxisX2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX2:Title:Font:Strikeout"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Strikeout
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX2:Title:Font:Underline"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Underline
                    SelectedChart.ChartAreas(AreaName).AxisX2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisX2:LabelStyleFormat"
                    SelectedChart.ChartAreas(AreaName).AxisX2.LabelStyle.Format = Data
                Case "ChartArea:AxisX2:Minimum"
                    SelectedChart.ChartAreas(AreaName).AxisX2.Minimum = Data
                Case "ChartArea:AxisX2:Maximum"
                    SelectedChart.ChartAreas(AreaName).AxisX2.Maximum = Data
                Case "ChartArea:AxisX2:LineWidth"
                    SelectedChart.ChartAreas(AreaName).AxisX2.LineWidth = Data
                Case "ChartArea:AxisX2:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisX2.Interval = Data
                Case "ChartArea:AxisX2:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisX2.IntervalOffset = Data
                Case "ChartArea:AxisX2:Crossing"
                    SelectedChart.ChartAreas(AreaName).AxisX2.Crossing = Data
                Case "ChartArea:AxisX2:MajorGrid:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisX2.MajorGrid.Interval = Data
                Case "ChartArea:AxisX2:MajorGrid:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisX2.MajorGrid.IntervalOffset = Data

            'Chart Area AxisY
                Case "ChartArea:AxisY:Title:Text"
                    SelectedChart.ChartAreas(AreaName).AxisY.Title = Data
                Case "ChartArea:AxisY:Title:Alignment"
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleAlignment = Data
                Case "ChartArea:AxisY:Title:ForeColor"
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleForeColor = Color.FromArgb(Data)
                Case "ChartArea:AxisY:Title:Font:Name"
                    ChartFontName = Data
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY:Title:Font:Size"
                    ChartFontSize = Data
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY:Title:Font:Bold"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Bold
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY:Title:Font:Italic"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Italic
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY:Title:Font:Strikeout"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Strikeout
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY:Title:Font:Underline"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Underline
                    SelectedChart.ChartAreas(AreaName).AxisY.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY:LabelStyleFormat"
                    SelectedChart.ChartAreas(AreaName).AxisY.LabelStyle.Format = Data
                Case "ChartArea:AxisY:Minimum"
                    SelectedChart.ChartAreas(AreaName).AxisY.Minimum = Data
                Case "ChartArea:AxisY:Maximum"
                    SelectedChart.ChartAreas(AreaName).AxisY.Maximum = Data
                Case "ChartArea:AxisY:LineWidth"
                    SelectedChart.ChartAreas(AreaName).AxisY.LineWidth = Data
                Case "ChartArea:AxisY:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisY.Interval = Data
                Case "ChartArea:AxisY:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisY.IntervalOffset = Data
                Case "ChartArea:AxisY:Crossing"
                    SelectedChart.ChartAreas(AreaName).AxisY.Crossing = Data
                Case "ChartArea:AxisY:MajorGrid:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisY.MajorGrid.Interval = Data
                Case "ChartArea:AxisY:MajorGrid:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisY.MajorGrid.IntervalOffset = Data

            'Chart Area AxisY2
                Case "ChartArea:AxisY2:Title:Text"
                    SelectedChart.ChartAreas(AreaName).AxisY2.Title = Data
                Case "ChartArea:AxisY2:Title:Alignment"
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleAlignment = Data
                Case "ChartArea:AxisY2:Title:ForeColor"
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleForeColor = Color.FromArgb(Data)
                Case "ChartArea:AxisY2:Title:Font:Name"
                    ChartFontName = Data
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY2:Title:Font:Size"
                    ChartFontSize = Data
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY2:Title:Font:Bold"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Bold
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY2:Title:Font:Italic"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Italic
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY2:Title:Font:Strikeout"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Strikeout
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY2:Title:Font:Underline"
                    ChartFontStyle = ChartFontStyle Or FontStyle.Underline
                    SelectedChart.ChartAreas(AreaName).AxisY2.TitleFont = New Font(ChartFontName, ChartFontSize, ChartFontStyle) 'Update the title font.
                Case "ChartArea:AxisY2:LabelStyleFormat"
                    SelectedChart.ChartAreas(AreaName).AxisY2.LabelStyle.Format = Data
                Case "ChartArea:AxisY2:Minimum"
                    SelectedChart.ChartAreas(AreaName).AxisY2.Minimum = Data
                Case "ChartArea:AxisY2:Maximum"
                    SelectedChart.ChartAreas(AreaName).AxisY2.Maximum = Data
                Case "ChartArea:AxisY2:LineWidth"
                    SelectedChart.ChartAreas(AreaName).AxisY2.LineWidth = Data
                Case "ChartArea:AxisY2:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisY2.Interval = Data
                Case "ChartArea:AxisY2:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisY2.IntervalOffset = Data
                Case "ChartArea:AxisY2:Crossing"
                    SelectedChart.ChartAreas(AreaName).AxisY2.Crossing = Data
                Case "ChartArea:AxisY2:MajorGrid:Interval"
                    SelectedChart.ChartAreas(AreaName).AxisY2.MajorGrid.Interval = Data
                Case "ChartArea:AxisY2:MajorGrid:IntervalOffset"
                    SelectedChart.ChartAreas(AreaName).AxisY2.MajorGrid.IntervalOffset = Data





            'End Of Stock Chart Instructions ----------------------------------------------------------------------------------------------------------------------------


            'Startup Command Arguments ================================================
                Case "ProjectName"
                    If Project.OpenProject(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                    Else
                        ProjectSelected = False 'Project could not be opened.
                    End If

                Case "ProjectID"
                    Message.AddWarning("Add code to handle ProjectID parameter at StartUp!" & vbCrLf)
                'Note the ComNet will usually select a project using ProjectPath.

                'Case "ProjectPath"
                '    If Project.OpenProjectPath(Data) = True Then
                '        ProjectSelected = True 'Project has been opened OK.
                '    Else
                '        ProjectSelected = False 'Project could not be opened.
                '    End If

                Case "ProjectPath"
                    If Project.OpenProjectPath(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                        'THE PROJECT IS LOCKED IN THE Form.Load EVENT:

                        ApplicationInfo.SettingsLocn = Project.SettingsLocn
                        Message.SettingsLocn = Project.SettingsLocn 'Set up the Message object
                        Message.Show() 'Added 18May19

                        'txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                        '              Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                        '              Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                        '              Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

                        'txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                        '               Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                        '               Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                        '               Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

                        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                                        Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                                        Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                                        Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

                        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & "d:" &
                                       Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & "h:" &
                                       Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & "m:" &
                                       Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c) & "s"

                    Else
                        ProjectSelected = False 'Project could not be opened.
                        Message.AddWarning("Project could not be opened at path: " & Data & vbCrLf)
                    End If

                Case "ConnectionName"
                    StartupConnectionName = Data
            '--------------------------------------------------------------------------

            'Application Information  =================================================
            'returned by client.GetAdvlNetworkAppInfoAsync()
            'Case "MessageServiceAppInfo:Name"
            '    'The name of the Message Service Application. (Not used.)
                Case "AdvlNetworkAppInfo:Name"
                'The name of the Andorville™ Network Application. (Not used.)

            'Case "MessageServiceAppInfo:ExePath"
            '    'The executable file path of the Message Service Application.
            '    MsgServiceExePath = Data
                Case "AdvlNetworkAppInfo:ExePath"
                    'The executable file path of the Andorville™ Network Application.
                    AdvlNetworkExePath = Data

            'Case "MessageServiceAppInfo:Path"
            '    'The path of the Message Service Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
            '    MsgServiceAppPath = Data
                Case "AdvlNetworkAppInfo:Path"
                    'The path of the Andorville™ Network Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                    AdvlNetworkAppPath = Data

           '---------------------------------------------------------------------------

            'Message Window Instructions  ==============================================
                Case "MessageWindow:Left"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Left = Data
                Case "MessageWindow:Top"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Top = Data
                Case "MessageWindow:Width"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Width = Data
                Case "MessageWindow:Height"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Height = Data
                Case "MessageWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            If IsNothing(Message.MessageForm) Then
                                Message.ApplicationName = ApplicationInfo.Name
                                Message.SettingsLocn = Project.SettingsLocn
                                Message.Show()
                            End If
                            'Message.MessageForm.BringToFront()
                            Message.MessageForm.Activate()
                            Message.MessageForm.TopMost = True
                            Message.MessageForm.TopMost = False
                        Case "SaveSettings"
                            Message.MessageForm.SaveFormSettings()
                    End Select

            '---------------------------------------------------------------------------

           'Command to bring the Application window to the front:
                Case "ApplicationWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            Me.Activate()
                            Me.TopMost = True
                            Me.TopMost = False
                    End Select


                Case "EndOfSequence"
                    'End of Information Vector Sequence reached.
                    'Add Status OK element at the end of the sequence:
                    Dim statusOK As New XElement("Status", "OK")
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    Select Case EndInstruction
                        Case "Stop"
                            'No instructions.

                            'Add any other Cases here:

                        Case Else
                            Message.AddWarning("Unknown End Instruction: " & EndInstruction & vbCrLf)
                    End Select
                    EndInstruction = "Stop"

                    ''Add the final OnCompletion instruction:
                    'Dim onCompletion As New XElement("OnCompletion", CompletionInstruction) '
                    'xlocns(xlocns.Count - 1).Add(onCompletion)
                    'CompletionInstruction = "Stop" 'Reset the Completion Instruction

                    'Add the final EndInstruction:
                    If OnCompletionInstruction = "Stop" Then
                        'Final EndInstruction is not required.
                    Else
                        Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                        xlocns(xlocns.Count - 1).Add(xEndInstruction)
                        OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction
                    End If

                Case Else
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            data: " & Data & vbCrLf)
            End Select
        End If
    End Sub

    Private Sub GetStockChartSettings(ByVal ClientLocn As String)
        'Return the current Stock Chart settings data in an XMsgBlk (XMessageBlock) file.

        If ConnectedToComNet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim ChartSettings As System.Xml.Linq.XDocument = ChartInfo.ToXDoc(Chart1)

                    Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                    'Dim xDataMsgDoc As New XDocument(Decl, Nothing)
                    Dim xMsgBlkDoc As New XDocument(Decl, Nothing)
                    'Dim xDataMsg As New XElement("XDataMsg")
                    Dim xMsgBlk As New XElement("XMsgBlk")

                    Dim clientLocation As New XElement("ClientLocn", ClientLocn)
                    'xDataMsg.Add(clientLocation)
                    xMsgBlk.Add(clientLocation)

                    'Dim xData As New XElement("XData")
                    Dim xInfo As New XElement("XInfo")
                    'Dim myData As XElement = ChartSettings.<ChartSettings> 'ERROR HERE
                    'Dim myData As Xml.XmlDocumentFragment = ChartSettings.<ChartSettings> 'ERROR HERE
                    'Dim myData As XElement = System.Xml.Linq.XElement.Parse(ChartSettings.<ChartSettings>.ToString)  'ERROR HERE
                    'xData.Add(myData)
                    'xData.Add(ChartSettings.<ChartSettings>)
                    xInfo.Add(ChartSettings.<ChartSettings>)
                    'xDataMsg.Add(xData)
                    'xMsgBlk.Add(xData)
                    xMsgBlk.Add(xInfo)

                    'xDataMsgDoc.Add(xDataMsg)
                    xMsgBlkDoc.Add(xMsgBlk)

                    'Show the message sent to the Message Service:
                    'Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")
                    'Message.XAddXml(xDataMsgDoc.ToString)
                    Message.XAddXml(xMsgBlkDoc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    'client.SendMessage("", "MessageService", xDataMsgDoc.ToString)
                    'client.SendMessage(ClientProNetName, ClientConnName, xDataMsgDoc.ToString) 'Error running XMsg: This operation would deadlock because the reply cannot be received until the current Message completes processing. If you want to allow out-of-order message processing, specify ConcurrencyMode of Reentrant or Multiple on CallbackBehaviorAttribute.

                    'Because Instructiosn are still being processed, use Alternative SendMessage background worker
                    SendMessageParamsAlt.ProjectNetworkName = ClientProNetName
                    SendMessageParamsAlt.ConnectionName = ClientConnName
                    'SendMessageParamsAlt.Message = xDataMsgDoc.ToString
                    SendMessageParamsAlt.Message = xMsgBlkDoc.ToString
                    If bgwSendMessageAlt.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        bgwSendMessageAlt.RunWorkerAsync(SendMessageParamsAlt)
                    End If
                End If
            End If
        End If
    End Sub

    'Private Sub SendMessage()
    '    'Code used to send a message after a timer delay.
    '    'The message destination is stored in MessageDest
    '    'The message text is stored in MessageText
    '    Timer1.Interval = 100 '100ms delay
    '    Timer1.Enabled = True 'Start the timer.
    'End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    '    If IsNothing(client) Then
    '        Message.AddWarning("No client connection available!" & vbCrLf)
    '    Else
    '        If client.State = ServiceModel.CommunicationState.Faulted Then
    '            Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
    '        Else
    '            Try
    '                'client.SendMessage(ClientAppNetName, ClientConnName, MessageText) 'Added 2Feb19
    '                client.SendMessage(ClientProNetName, ClientConnName, MessageText)
    '                MessageText = "" 'Clear the message after it has been sent.
    '                ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
    '                ClientConnName = "" 'Clear the Client Application Name after the message has been sent.
    '                xlocns.Clear()
    '            Catch ex As Exception
    '                Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
    '            End Try
    '        End If
    '    End If

    '    'Stop timer:
    '    Timer1.Enabled = False
    'End Sub

    'Private Sub GetStockChartSettingsForClient()
    '    'Get the Stock Chart settings and send it to the Client.

    '    Dim chartSettings As New XElement("Settings")

    '    Dim commandClearChart As New XElement("Command", "ClearChart")
    '    chartSettings.Add(commandClearChart)

    '    Dim inputData As New XElement("InputData")

    '    Dim dataType As New XElement("Type", StockChart.InputDataType)
    '    inputData.Add(dataType)

    '    If Trim(StockChart.InputDatabasePath) = "" Then
    '        Dim databasePath As New XElement("DatabasePath", "-")
    '        inputData.Add(databasePath)
    '    Else
    '        Dim databasePath As New XElement("DatabasePath", StockChart.InputDatabasePath)
    '        inputData.Add(databasePath)
    '    End If

    '    If Trim(StockChart.InputDataDescr) = "" Then
    '        Dim dataDescription As New XElement("DataDescription", "-")
    '        inputData.Add(dataDescription)
    '    Else
    '        Dim dataDescription As New XElement("DataDescription", StockChart.InputDataDescr)
    '        inputData.Add(dataDescription)
    '    End If

    '    If Trim(StockChart.InputQuery) = "" Then
    '        Dim databaseQuery As New XElement("DatabaseQuery", "-")
    '        inputData.Add(databaseQuery)
    '    Else
    '        Dim databaseQuery As New XElement("DatabaseQuery", StockChart.InputQuery)
    '        inputData.Add(databaseQuery)
    '    End If

    '    chartSettings.Add(inputData)

    '    Dim chartProperties As New XElement("ChartProperties")

    '    If Trim(StockChart.SeriesName) = "" Then
    '        Dim seriesName As New XElement("SeriesName", "-")
    '        chartProperties.Add(seriesName)
    '    Else
    '        Dim seriesName As New XElement("SeriesName", StockChart.SeriesName)
    '        chartProperties.Add(seriesName)
    '    End If

    '    If Trim(StockChart.XValuesFieldName) = "" Then
    '        Dim xValuesFieldName As New XElement("XValuesFieldName", "-")
    '        chartProperties.Add(xValuesFieldName)
    '    Else
    '        Dim xValuesFieldName As New XElement("XValuesFieldName", StockChart.XValuesFieldName)
    '        chartProperties.Add(xValuesFieldName)
    '    End If

    '    If Trim(StockChart.YValuesHighFieldName) = "" Then
    '        Dim yValuesHighFieldName As New XElement("YValuesHighFieldName", "-")
    '        chartProperties.Add(yValuesHighFieldName)
    '    Else
    '        Dim yValuesHighFieldName As New XElement("YValuesHighFieldName", StockChart.YValuesHighFieldName)
    '        chartProperties.Add(yValuesHighFieldName)
    '    End If

    '    If Trim(StockChart.YValuesLowFieldName) = "" Then
    '        Dim yValuesLowFieldName As New XElement("YValuesLowFieldName", "-")
    '        chartProperties.Add(yValuesLowFieldName)
    '    Else
    '        Dim yValuesLowFieldName As New XElement("YValuesLowFieldName", StockChart.YValuesLowFieldName)
    '        chartProperties.Add(yValuesLowFieldName)
    '    End If

    '    If Trim(StockChart.YValuesOpenFieldName) = "" Then
    '        Dim yValuesOpenFieldName As New XElement("YValuesOpenFieldName", "-")
    '        chartProperties.Add(yValuesOpenFieldName)
    '    Else
    '        Dim yValuesOpenFieldName As New XElement("YValuesOpenFieldName", StockChart.YValuesOpenFieldName)
    '        chartProperties.Add(yValuesOpenFieldName)
    '    End If

    '    Dim yValuesCloseFieldName As New XElement("YValuesCloseFieldName", StockChart.YValuesCloseFieldName)
    '    chartProperties.Add(yValuesCloseFieldName)
    '    chartSettings.Add(chartProperties)

    '    Dim chartTitle As New XElement("ChartTitle")
    '    If Trim(StockChart.ChartLabel.Name) = "" Then
    '        Dim chartTitleLabelName As New XElement("LabelName", "Label1")
    '        chartTitle.Add(chartTitleLabelName)
    '    Else
    '        Dim chartTitleLabelName As New XElement("LabelName", StockChart.ChartLabel.Name)
    '        chartTitle.Add(chartTitleLabelName)
    '    End If

    '    If Trim(StockChart.ChartLabel.Text) = "" Then
    '        Dim chartTitleText As New XElement("Text", "-")
    '        chartTitle.Add(chartTitleText)
    '    Else
    '        Dim chartTitleText As New XElement("Text", StockChart.ChartLabel.Text)
    '        chartTitle.Add(chartTitleText)
    '    End If

    '    Dim chartTitleFontName As New XElement("FontName", StockChart.ChartLabel.FontName)
    '    chartTitle.Add(chartTitleFontName)
    '    'Dim chartTitleColor As New XElement("Color", StockChart.ChartLabel.Color)
    '    Dim chartTitleColor As New XElement("Color", StockChart.ChartLabel.Color.ToArgb.ToString)
    '    chartTitle.Add(chartTitleColor)
    '    Dim chartTitleSize As New XElement("Size", StockChart.ChartLabel.Size)
    '    chartTitle.Add(chartTitleSize)
    '    Dim chartTitleBold As New XElement("Bold", StockChart.ChartLabel.Bold)
    '    chartTitle.Add(chartTitleBold)
    '    Dim chartTitleItalic As New XElement("Italic", StockChart.ChartLabel.Italic)
    '    chartTitle.Add(chartTitleItalic)
    '    Dim chartTitleUnderline As New XElement("Underline", StockChart.ChartLabel.Underline)
    '    chartTitle.Add(chartTitleUnderline)
    '    Dim chartTitleStrikeout As New XElement("Strikeout", StockChart.ChartLabel.Strikeout)
    '    chartTitle.Add(chartTitleStrikeout)
    '    Dim chartTitleAlignment As New XElement("Alignment", StockChart.ChartLabel.Alignment)
    '    chartTitle.Add(chartTitleAlignment)

    '    chartSettings.Add(chartTitle)

    '    Dim xAxis As New XElement("XAxis")
    '    If Trim(StockChart.XAxis.Title.Text) = "" Then
    '        Dim titleText As New XElement("TitleText", "-")
    '        xAxis.Add(titleText)
    '    Else
    '        Dim titleText As New XElement("TitleText", StockChart.XAxis.Title.Text)
    '        xAxis.Add(titleText)
    '    End If

    '    Dim titleFontName As New XElement("TitleFontName", StockChart.XAxis.Title.FontName)
    '    xAxis.Add(titleFontName)
    '    Dim titleFontColor As New XElement("TitleColor", StockChart.XAxis.Title.Color) 'This is stred as a string!
    '    xAxis.Add(titleFontColor)
    '    Dim titleSize As New XElement("TitleSize", StockChart.XAxis.Title.Size)
    '    xAxis.Add(titleSize)
    '    Dim titleBold As New XElement("TitleBold", StockChart.XAxis.Title.Bold)
    '    xAxis.Add(titleBold)
    '    Dim titleItalic As New XElement("TitleItalic", StockChart.XAxis.Title.Italic)
    '    xAxis.Add(titleItalic)
    '    Dim titleUnderline As New XElement("TitleUnderline", StockChart.XAxis.Title.Underline)
    '    xAxis.Add(titleUnderline)
    '    Dim titleStrikeout As New XElement("TitleStrikeout", StockChart.XAxis.Title.Strikeout)
    '    xAxis.Add(titleStrikeout)
    '    Dim titleAlignment As New XElement("TitleAlignment", StockChart.XAxis.TitleAlignment)
    '    xAxis.Add(titleAlignment)
    '    Dim autoMinimum As New XElement("AutoMinimum", StockChart.XAxis.AutoMinimum)
    '    xAxis.Add(autoMinimum)
    '    Dim minimum As New XElement("Minimum", StockChart.XAxis.Minimum)
    '    xAxis.Add(minimum)
    '    Dim autoMaximum As New XElement("AutoMaximum", StockChart.XAxis.AutoMaximum)
    '    xAxis.Add(autoMaximum)
    '    Dim maximum As New XElement("Maximum", StockChart.XAxis.Maximum)
    '    xAxis.Add(maximum)
    '    Dim autoInterval As New XElement("AutoInterval", StockChart.XAxis.AutoInterval)
    '    xAxis.Add(autoInterval)
    '    Dim interval As New XElement("Interval", StockChart.XAxis.Interval)
    '    xAxis.Add(interval)
    '    Dim autoMajorGridInterval As New XElement("AutoMajorGridInterval", StockChart.XAxis.AutoMajorGridInterval)
    '    xAxis.Add(autoMajorGridInterval)
    '    Dim majorGridInterval As New XElement("MajorGridInterval", StockChart.XAxis.MajorGridInterval)
    '    xAxis.Add(majorGridInterval)
    '    chartSettings.Add(xAxis)

    '    Dim yAxis As New XElement("YAxis")
    '    If Trim(StockChart.YAxis.Title.Text) = "" Then
    '        Dim titleText2 As New XElement("TitleText", "-")
    '        yAxis.Add(titleText2)
    '    Else
    '        Dim titleText2 As New XElement("TitleText", StockChart.YAxis.Title.Text)
    '        yAxis.Add(titleText2)
    '    End If

    '    Dim titleFontName2 As New XElement("TitleFontName", StockChart.YAxis.Title.FontName)
    '    yAxis.Add(titleFontName2)
    '    Dim titleFontColor2 As New XElement("TitleColor", StockChart.YAxis.Title.Color) 'This is stored as a string.
    '    yAxis.Add(titleFontColor2)
    '    Dim titleSize2 As New XElement("TitleSize", StockChart.YAxis.Title.Size)
    '    yAxis.Add(titleSize2)
    '    Dim titleBold2 As New XElement("TitleBold", StockChart.YAxis.Title.Bold)
    '    yAxis.Add(titleBold2)
    '    Dim titleItalic2 As New XElement("TitleItalic", StockChart.YAxis.Title.Italic)
    '    yAxis.Add(titleItalic2)
    '    Dim titleUnderline2 As New XElement("TitleUnderline", StockChart.YAxis.Title.Underline)
    '    yAxis.Add(titleUnderline2)
    '    Dim titleStrikeout2 As New XElement("TitleStrikeout", StockChart.YAxis.Title.Strikeout)
    '    yAxis.Add(titleStrikeout2)
    '    Dim titleAlignment2 As New XElement("TitleAlignment", StockChart.YAxis.TitleAlignment)
    '    yAxis.Add(titleAlignment2)
    '    Dim autoMinimum2 As New XElement("AutoMinimum", StockChart.YAxis.AutoMinimum)
    '    yAxis.Add(autoMinimum2)
    '    Dim minimum2 As New XElement("Minimum", StockChart.YAxis.Minimum)
    '    yAxis.Add(minimum2)
    '    Dim autoMaximum2 As New XElement("AutoMaximum", StockChart.YAxis.AutoMaximum)
    '    yAxis.Add(autoMaximum2)
    '    Dim maximum2 As New XElement("Maximum", StockChart.YAxis.Maximum)
    '    yAxis.Add(maximum2)
    '    Dim autoInterval2 As New XElement("AutoInterval", StockChart.YAxis.AutoInterval)
    '    yAxis.Add(autoInterval2)
    '    Dim interval2 As New XElement("Interval", StockChart.YAxis.Interval)
    '    yAxis.Add(interval2)
    '    Dim autoMajorGridInterval2 As New XElement("AutoMajorGridInterval", StockChart.YAxis.AutoMajorGridInterval)
    '    yAxis.Add(autoMajorGridInterval2)
    '    Dim majorGridInterval2 As New XElement("MajorGridInterval", StockChart.YAxis.MajorGridInterval)
    '    yAxis.Add(majorGridInterval2)
    '    chartSettings.Add(yAxis)

    '    Dim commandOK As New XElement("Command", "OK")
    '    chartSettings.Add(commandOK)

    '    xlocns(xlocns.Count - 1).Add(chartSettings) 'The settings are aded to the last location in the XLocations list.

    'End Sub


    'Private Sub btnDatabase_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
    '    'Select a database file:

    '    OpenFileDialog1.Filter = "Access Database |*.accdb"
    '    OpenFileDialog1.FileName = ""

    '    If InputDatabaseDirectory <> "" Then
    '        OpenFileDialog1.InitialDirectory = InputDatabaseDirectory
    '    End If

    '    OpenFileDialog1.ShowDialog()

    '    StockChart.InputDatabasePath = OpenFileDialog1.FileName
    '    txtDatabasePath.Text = StockChart.InputDatabasePath
    '    'FillLstTables()

    '    InputDatabaseDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

    '    Message.Add("InputDatabaseDirectory = " & InputDatabaseDirectory & vbCrLf)

    '    FillLstTables()

    'End Sub

    'Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
    '    StockChart.InputQuery = txtInputQuery.Text
    '    ApplyQuery() 'This fills the Dataset with the Query data.
    '    GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
    '    UpdateStockChartForm()
    'End Sub


    'Private Sub SetUpStockChartStyleTab()
    '    'Set up the Stock Chart style tab.

    '    txtChartDescr.Text = txtChartDescr.Text & "A Stock chart is typically used to illustrate significant stock price points including a stock's open, close, high, and low price points. However, this type of chart can also be used to analyze scientific data, because each series of data displays high, low, open, and close values, which are typically lines or triangles. The opening values are shown by the markers on the left, and the closing values are shown by the markers on the right." & vbCrLf

    '    'Initialise Chart Settings Tab ------------------------------------------------------
    '    'Set up the Y Values grid:
    '    DataGridView1.ColumnCount = 1
    '    DataGridView1.RowCount = 1
    '    DataGridView1.Columns(0).HeaderText = "Y Value"
    '    DataGridView1.Columns(0).Width = 60
    '    DataGridView1.Columns.Insert(1, cboFieldSelections)
    '    DataGridView1.Columns(1).HeaderText = "Field"
    '    DataGridView1.Columns(1).Width = 120
    '    DataGridView1.AllowUserToResizeColumns = True

    '    'Set up the Custom Attributes grid:
    '    'DataGridView2.ColumnCount = 3
    '    DataGridView2.ColumnCount = 4
    '    DataGridView2.RowCount = 1
    '    DataGridView2.Columns(0).HeaderText = "Custom Attribute"
    '    DataGridView2.Columns(0).Width = 120
    '    DataGridView2.Columns(1).HeaderText = "Value Range"
    '    DataGridView2.Columns(1).Width = 120
    '    DataGridView2.Columns(2).HeaderText = "Value"
    '    DataGridView2.Columns(2).Width = 120
    '    DataGridView2.Columns(3).HeaderText = "Description"
    '    DataGridView2.Columns(3).Width = 340
    '    DataGridView2.AllowUserToResizeColumns = True


    '    'Y Values:
    '    DataGridView1.Rows.Clear()
    '    DataGridView1.Rows.Add(4)
    '    DataGridView1.Rows(0).Cells(0).Value = "High" 'First Y Value Parameter Name
    '    DataGridView1.Rows(1).Cells(0).Value = "Low" 'Second Y Value parameter name
    '    DataGridView1.Rows(2).Cells(0).Value = "Open" 'Third Y Value parameter name
    '    DataGridView1.Rows(3).Cells(0).Value = "Close" 'Fourth Y value parameter name

    '    'Custom Attributes:
    '    DataGridView2.Rows.Clear()
    '    DataGridView2.Rows.Add(9)
    '    DataGridView2.Rows(0).Cells(0).Value = "LabelValueType"
    '    DataGridView2.Rows(0).Cells(1).Value = "High, Low, Open, Close"
    '    Dim cbc0 As New DataGridViewComboBoxCell
    '    cbc0.Items.Add(" ")
    '    cbc0.Items.Add("High")
    '    cbc0.Items.Add("Low")
    '    cbc0.Items.Add("Open")
    '    cbc0.Items.Add("Close")
    '    DataGridView2.Rows(0).Cells(2) = cbc0
    '    DataGridView2.Rows(0).Cells(3).Value = "Specifies the Y value to use as the data point label."
    '    DataGridView2.Rows(1).Cells(0).Value = "MaxPixelPointWidth"
    '    DataGridView2.Rows(1).Cells(1).Value = "Any integer > 0"
    '    DataGridView2.Rows(1).Cells(3).Value = "Specifies the maximum width of the data point in pixels."
    '    DataGridView2.Rows(2).Cells(0).Value = "MinPixelPointWidth"
    '    DataGridView2.Rows(2).Cells(1).Value = "Any integer > 0"
    '    DataGridView2.Rows(2).Cells(3).Value = "Specifies the minimum data point width in pixels."
    '    DataGridView2.Rows(3).Cells(0).Value = "OpenCloseStyle"
    '    DataGridView2.Rows(3).Cells(1).Value = "Triangle, Line, Candlestick"
    '    Dim cbc3 As New DataGridViewComboBoxCell
    '    cbc3.Items.Add(" ")
    '    cbc3.Items.Add("Triangle")
    '    cbc3.Items.Add("Line")
    '    cbc3.Items.Add("Candlestick")
    '    DataGridView2.Rows(3).Cells(2) = cbc3
    '    DataGridView2.Rows(3).Cells(3).Value = "Specifies the marker style for open and close values."
    '    DataGridView2.Rows(4).Cells(0).Value = "PixelPointDepth"
    '    DataGridView2.Rows(4).Cells(1).Value = "Any integer > 0"
    '    DataGridView2.Rows(4).Cells(3).Value = "Specifies the 3D series depth in pixels."
    '    DataGridView2.Rows(5).Cells(0).Value = "PixelPointGapDepth"
    '    DataGridView2.Rows(5).Cells(1).Value = "Any integer > 0"
    '    DataGridView2.Rows(5).Cells(3).Value = "Specifies the 3D gap depth in pixels."
    '    DataGridView2.Rows(6).Cells(0).Value = "PixelPointWidth"
    '    DataGridView2.Rows(6).Cells(1).Value = "Any integer > 0"
    '    DataGridView2.Rows(6).Cells(3).Value = "Specifies the data point width in pixels."
    '    DataGridView2.Rows(7).Cells(0).Value = "PointWidth"
    '    DataGridView2.Rows(7).Cells(1).Value = "0 to 2"
    '    DataGridView2.Rows(7).Cells(3).Value = "Specifies data point width."
    '    DataGridView2.Rows(8).Cells(0).Value = "ShowOpenClose"
    '    DataGridView2.Rows(8).Cells(1).Value = "Both, Open, Close"
    '    DataGridView2.Rows(8).Cells(3).Value = "Specifies whether markers for open and close prices are displayed."
    '    Dim cbc8 As New DataGridViewComboBoxCell
    '    cbc8.Items.Add(" ")
    '    cbc8.Items.Add("Both")
    '    cbc8.Items.Add("Open")
    '    cbc8.Items.Add("Close")
    '    DataGridView2.Rows(8).Cells(2) = cbc8

    'End Sub

    Private Sub UpdateStockChartTab()

    End Sub

    'Private Sub GetChartOptionsFromDataset()
    '    'Update the Chart display options from the Dataset.

    '    'First check that data to chart has been loaded into the dataset:
    '    If ds.Tables.Count = 0 Then
    '        Message.AddWarning("No data has been selected for charting." & vbCrLf)
    '    Else
    '        'Get the list of available fields from the dataset:
    '        GetFieldListFromDataset()
    '    End If

    '    'Show the selected XValues field: --------------------------------------------------------------------------------
    '    Dim I As Integer 'Loop index
    '    For I = 1 To cmbXValues.Items.Count
    '        If cmbXValues.Items(I - 1) = StockChart.XValuesFieldName Then
    '            cmbXValues.SelectedIndex = I - 1
    '        End If
    '    Next

    'End Sub


    'Private Sub GetFieldListFromDataset()
    '    'Update the available list of fields for plotting on the X and Y axes.

    '    cboFieldSelections.Items.Clear()
    '    cmbXValues.Items.Clear()

    '    If ds.Tables(0).Columns.Count > 0 Then
    '        Dim I As Integer 'Loop index
    '        For I = 1 To ds.Tables(0).Columns.Count
    '            cboFieldSelections.Items.Add(ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox column used in DataGridView1 (Y Values to chart)
    '            cmbXValues.Items.Add(ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox used to select the Field to use along the X Axis
    '        Next
    '    End If
    'End Sub

    'Private Sub UpdateStockChartForm()
    '    'Update the Chart Type, Titles, X Axis and Y Axis tabs with the settings stored in StockChart.

    '    txtDatabasePath.Text = StockChart.InputDatabasePath
    '    txtDataDescription.Text = StockChart.InputDataDescr
    '    txtInputQuery.Text = StockChart.InputQuery
    '    txtDataDescription.Text = StockChart.InputDataDescr


    '    txtSeriesName.Text = StockChart.SeriesName

    '    'Apply X Values Field Name:
    '    cmbXValues.SelectedIndex = cmbXValues.FindStringExact(StockChart.XValuesFieldName)

    '    'Apply Y Values settings:
    '    For I = 1 To cboFieldSelections.Items.Count
    '        If StockChart.YValuesHighFieldName = cboFieldSelections.Items(I - 1) Then
    '            DataGridView1.Rows(0).Cells(1).Value = cboFieldSelections.Items(I - 1)
    '        End If
    '        If StockChart.YValuesLowFieldName = cboFieldSelections.Items(I - 1) Then
    '            DataGridView1.Rows(1).Cells(1).Value = cboFieldSelections.Items(I - 1)
    '        End If
    '        If StockChart.YValuesOpenFieldName = cboFieldSelections.Items(I - 1) Then
    '            DataGridView1.Rows(2).Cells(1).Value = cboFieldSelections.Items(I - 1)
    '        End If
    '        If StockChart.YValuesCloseFieldName = cboFieldSelections.Items(I - 1) Then
    '            DataGridView1.Rows(3).Cells(1).Value = cboFieldSelections.Items(I - 1)
    '        End If
    '    Next
    '    'Apply Custom Attributes Settings:
    '    'LabelValueType (High, Low, Open, Close)
    '    If StockChart.LabelValueType <> "" Then
    '        DataGridView2.Rows(0).Cells(2).Value = StockChart.LabelValueType 'This produces an error
    '    End If
    '    'MaxPixelPointWidth (Any integer > 0)
    '    DataGridView2.Rows(1).Cells(2).Value = StockChart.MaxPixelPointWidth
    '    'MinPixelPointWidth (Any integer > 0)
    '    DataGridView2.Rows(2).Cells(2).Value = StockChart.MinPixelPointWidth
    '    'OpenCloseStyle (Triangle, Line, Candle)
    '    If StockChart.OpenCloseStyle <> "" Then
    '        DataGridView2.Rows(3).Cells(2).Value = StockChart.OpenCloseStyle
    '    End If
    '    'PixelPointDepth (Any integer > 0)
    '    DataGridView2.Rows(4).Cells(2).Value = StockChart.PixelPointDepth
    '    'PixelPointGapDepth (Any integer > 0)
    '    DataGridView2.Rows(5).Cells(2).Value = StockChart.PixelPointGapDepth
    '    'PixelPointWidth (Any integer > 0)
    '    DataGridView2.Rows(6).Cells(2).Value = StockChart.PixelPointWidth
    '    'PointWidth (0 to 2)
    '    DataGridView2.Rows(7).Cells(2).Value = StockChart.PointWidth
    '    'ShowOpenClose (Both, Open, Close)
    '    If StockChart.ShowOpenClose <> "" Then
    '        DataGridView2.Rows(8).Cells(2).Value = StockChart.ShowOpenClose
    '    End If

    '    'Update the ChartLabel settings: -------------------------------------------------------------------------
    '    txtChartTitle.Text = StockChart.ChartLabel.Text
    '    'txtChartTitle.ForeColor = Color.FromName(StockChart.ChartLabel.Color)
    '    txtChartTitle.ForeColor = StockChart.ChartLabel.Color

    '    Dim myFontStyle As FontStyle = FontStyle.Regular
    '    If StockChart.ChartLabel.Bold Then
    '        myFontStyle = myFontStyle Or FontStyle.Bold
    '    End If
    '    If StockChart.ChartLabel.Italic Then
    '        myFontStyle = myFontStyle Or FontStyle.Italic
    '    End If
    '    If StockChart.ChartLabel.Strikeout Then
    '        myFontStyle = myFontStyle Or FontStyle.Strikeout
    '    End If
    '    If StockChart.ChartLabel.Underline Then
    '        myFontStyle = myFontStyle Or FontStyle.Underline
    '    End If

    '    txtChartTitle.Font = New Font("Arial", StockChart.ChartLabel.Size, myFontStyle)

    '    'Update the XAxis settings: -------------------------------------------------------------------------
    '    txtXAxisTitle.Text = StockChart.XAxis.Title.Text
    '    txtXAxisTitle.ForeColor = Color.FromName(StockChart.XAxis.Title.Color)

    '    myFontStyle = FontStyle.Regular
    '    If StockChart.XAxis.Title.Bold Then
    '        myFontStyle = myFontStyle Or FontStyle.Bold
    '    End If
    '    If StockChart.XAxis.Title.Italic Then
    '        myFontStyle = myFontStyle Or FontStyle.Italic
    '    End If
    '    If StockChart.XAxis.Title.Strikeout Then
    '        myFontStyle = myFontStyle Or FontStyle.Strikeout
    '    End If
    '    If StockChart.XAxis.Title.Underline Then
    '        myFontStyle = myFontStyle Or FontStyle.Underline
    '    End If

    '    txtXAxisTitle.Font = New Font(StockChart.XAxis.Title.FontName, StockChart.XAxis.Title.Size, myFontStyle)

    '    chkXAxisAutoMin.Checked = StockChart.XAxis.AutoMinimum
    '    chkXAxisAutoMax.Checked = StockChart.XAxis.AutoMaximum

    '    txtXAxisMin.Text = StockChart.XAxis.Minimum
    '    txtXAxisMax.Text = StockChart.XAxis.Maximum

    '    chkXAxisAutoAnnotInt.Checked = StockChart.XAxis.AutoInterval
    '    chkXAxisAutoMajGridInt.Checked = StockChart.XAxis.AutoMajorGridInterval

    '    txtXAxisAnnotInt.Text = StockChart.XAxis.Interval
    '    txtXAxisMajGridInt.Text = StockChart.XAxis.MajorGridInterval

    '    'Update the YAxis settings: -----------------------------------------------------------------------------
    '    txtYAxisTitle.Text = StockChart.YAxis.Title.Text
    '    txtYAxisTitle.ForeColor = Color.FromName(StockChart.YAxis.Title.Color)

    '    myFontStyle = FontStyle.Regular
    '    If StockChart.YAxis.Title.Bold Then
    '        myFontStyle = myFontStyle Or FontStyle.Bold
    '    End If
    '    If StockChart.YAxis.Title.Italic Then
    '        myFontStyle = myFontStyle Or FontStyle.Italic
    '    End If
    '    If StockChart.YAxis.Title.Strikeout Then
    '        myFontStyle = myFontStyle Or FontStyle.Strikeout
    '    End If
    '    If StockChart.YAxis.Title.Underline Then
    '        myFontStyle = myFontStyle Or FontStyle.Underline
    '    End If

    '    txtYAxisTitle.Font = New Font(StockChart.YAxis.Title.FontName, StockChart.YAxis.Title.Size, myFontStyle)

    '    chkYAxisAutoMin.Checked = StockChart.YAxis.AutoMinimum
    '    chkYAxisAutoMax.Checked = StockChart.YAxis.AutoMaximum

    '    txtYAxisMin.Text = StockChart.YAxis.Minimum
    '    txtYAxisMax.Text = StockChart.YAxis.Maximum

    '    chkYAxisAutoAnnotInt.Checked = StockChart.YAxis.AutoInterval
    '    chkYAxisAutoMajGridInt.Checked = StockChart.YAxis.AutoMajorGridInterval

    '    txtYAxisAnnotInt.Text = StockChart.YAxis.Interval
    '    txtYAxisMajGridInt.Text = StockChart.YAxis.MajorGridInterval

    '    'Update chart File Name:
    '    txtChartFileName.Text = StockChart.FileName

    'End Sub

    Private Sub btnDrawChart_Click(sender As Object, e As EventArgs) Handles btnDrawChart.Click
        DrawStockChart()
    End Sub

    Private Sub DrawStockChart()
        'Draw the Stock Chart:
        Try
            Dim SeriesName As String
            Dim ChartArea As String
            For Each item In Chart1.Series
                SeriesName = item.Name
                ChartArea = Chart1.Series(SeriesName).ChartArea
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMinimum Then Chart1.ChartAreas(ChartArea).AxisX.Minimum = Double.NaN
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMaximum Then Chart1.ChartAreas(ChartArea).AxisX.Maximum = Double.NaN
                If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMajorGridInterval Then Chart1.ChartAreas(ChartArea).AxisX.MajorGrid.Interval = Double.NaN
                Chart1.ChartAreas(ChartArea).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
                If item.ChartType = DataVisualization.Charting.SeriesChartType.Stock Then
                    'Chart1.Series(SeriesName).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictFields(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictFields(SeriesName).YValuesFieldName)
                    'Chart1.Series(SeriesName).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName)
                    Chart1.Series(SeriesName).YValuesPerPoint = 4
                    Chart1.Series(SeriesName).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictSeriesInfo(SeriesName).YValuesHighFieldName & "," & ChartInfo.dictSeriesInfo(SeriesName).YValuesLowFieldName & "," & ChartInfo.dictSeriesInfo(SeriesName).YValuesOpenFieldName & "," & ChartInfo.dictSeriesInfo(SeriesName).YValuesCloseFieldName)
                End If
                'If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMinimum Then Chart1.ChartAreas(ChartArea).AxisX.Minimum = Double.NaN
                'If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMaximum Then Chart1.ChartAreas(ChartArea).AxisX.Maximum = Double.NaN
                'If ChartInfo.dictAreaInfo(ChartArea).AutoXAxisMajorGridInterval Then Chart1.ChartAreas(ChartArea).AxisX.MajorGrid.Interval = Double.NaN
                'Chart1.ChartAreas(ChartArea).AxisX.Interval = 0
                'Chart1.ChartAreas(ChartArea).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
            Next
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub

    'Private Sub DrawStockChart()
    '    'Draw the Stock Chart using the settings specified in StockChart

    '    'Draw the Stock Chart:

    '    UpdateStockChartSettings()
    '    ApplyQuery() 'This fills the Dataset with the Query data.

    '    Try
    '        Chart1.Series.Clear()
    '        Chart1.Series.Add(StockChart.SeriesName)
    '        Chart1.Series(StockChart.SeriesName).YValuesPerPoint = 4
    '        Chart1.Series(StockChart.SeriesName).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, StockChart.XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, StockChart.YValuesHighFieldName & "," & StockChart.YValuesLowFieldName & "," & StockChart.YValuesOpenFieldName & "," & StockChart.YValuesCloseFieldName)
    '        Chart1.Series(StockChart.SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Stock
    '        If StockChart.LabelValueType <> "" Then
    '            Chart1.Series(StockChart.SeriesName).SetCustomProperty("LabelValueType", StockChart.LabelValueType)
    '        End If

    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("MaxPixelPointWidth", StockChart.MaxPixelPointWidth)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("MinPixelPointWidth", StockChart.MinPixelPointWidth)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("OpenCloseStyle", StockChart.OpenCloseStyle)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("PixelPointDepth", StockChart.PixelPointDepth)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("PixelPointGapDepth", StockChart.PixelPointGapDepth)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("PixelPointWidth", StockChart.PixelPointWidth)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("PointWidth", StockChart.PointWidth)
    '        Chart1.Series(StockChart.SeriesName).SetCustomProperty("ShowOpenClose", StockChart.ShowOpenClose)

    '        'Specify Y Axis range: -------------------------------------------------------------------------------
    '        If StockChart.YAxis.AutoMinimum = True Then
    '            Chart1.ChartAreas(0).AxisY.Minimum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisY.Minimum = StockChart.YAxis.Minimum
    '        End If
    '        If StockChart.YAxis.AutoMaximum = True Then
    '            Chart1.ChartAreas(0).AxisY.Maximum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisY.Maximum = StockChart.YAxis.Maximum
    '        End If

    '        'Specify Y Axis annotation and major grid intervals: -----------------------------------------------------
    '        If StockChart.YAxis.AutoInterval = True Then
    '            Chart1.ChartAreas(0).AxisY.Interval = 0
    '        Else
    '            Chart1.ChartAreas(0).AxisY.Interval = StockChart.YAxis.Interval
    '        End If

    '        If StockChart.YAxis.AutoMajorGridInterval = True Then
    '            Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = 0
    '            'Message.Add("Y Axis major grid interval is automatic." & vbCrLf)
    '        Else
    '            Chart1.ChartAreas(0).AxisY.MajorGrid.Interval = StockChart.YAxis.MajorGridInterval
    '        End If


    '        'Specify X Axis range: ------------------------------------------------------------------------------
    '        If StockChart.XAxis.AutoMinimum = True Then
    '            Chart1.ChartAreas(0).AxisX.Minimum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisX.Minimum = StockChart.XAxis.Minimum
    '        End If
    '        If StockChart.XAxis.AutoMaximum = True Then
    '            Chart1.ChartAreas(0).AxisX.Maximum = [Double].NaN
    '        Else
    '            Chart1.ChartAreas(0).AxisX.Maximum = StockChart.XAxis.Maximum
    '        End If

    '        'Specify X Axis annotation and major grid intervals: -----------------------------------------------------
    '        Chart1.ChartAreas(0).AxisX.IntervalType = Charting.DateTimeIntervalType.Auto

    '        If StockChart.XAxis.AutoInterval = True Then
    '            Chart1.ChartAreas(0).AxisX.Interval = 0
    '        Else
    '            Chart1.ChartAreas(0).AxisX.Interval = StockChart.XAxis.Interval
    '        End If

    '        If StockChart.XAxis.AutoMajorGridInterval = True Then
    '            Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = 0
    '            'Message.Add("X Axis major grid interval is automatic." & vbCrLf)
    '        Else
    '            Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = StockChart.XAxis.MajorGridInterval
    '        End If

    '        'Chart1.ChartAreas(0).RecalculateAxesScale()

    '        Chart1.ChartAreas(0).AxisX.LabelStyle.IsEndLabelVisible = True

    '        'Specify X Axis label: ------------------------------------------------------------------------------------
    '        Chart1.ChartAreas(0).AxisX.TitleAlignment = StockChart.XAxis.TitleAlignment

    '        Dim myFontStyle As FontStyle = FontStyle.Regular
    '        If StockChart.XAxis.Title.Bold Then
    '            myFontStyle = myFontStyle Or FontStyle.Bold
    '        End If
    '        If StockChart.XAxis.Title.Italic Then
    '            myFontStyle = myFontStyle Or FontStyle.Italic
    '        End If
    '        If StockChart.XAxis.Title.Strikeout Then
    '            myFontStyle = myFontStyle Or FontStyle.Strikeout
    '        End If
    '        If StockChart.XAxis.Title.Underline Then
    '            myFontStyle = myFontStyle Or FontStyle.Underline
    '        End If

    '        Chart1.ChartAreas(0).AxisX.TitleFont = New Font("Arial", StockChart.XAxis.Title.Size, myFontStyle)
    '        Chart1.ChartAreas(0).AxisX.Title = StockChart.XAxis.Title.Text

    '        'Specify Y Axis label: ------------------------------------------------------------------------------------
    '        Chart1.ChartAreas(0).AxisY.TitleAlignment = StockChart.YAxis.TitleAlignment
    '        myFontStyle = FontStyle.Regular
    '        If StockChart.YAxis.Title.Bold Then
    '            myFontStyle = myFontStyle Or FontStyle.Bold
    '        End If
    '        If StockChart.YAxis.Title.Italic Then
    '            myFontStyle = myFontStyle Or FontStyle.Italic
    '        End If
    '        If StockChart.YAxis.Title.Strikeout Then
    '            myFontStyle = myFontStyle Or FontStyle.Strikeout
    '        End If
    '        If StockChart.YAxis.Title.Underline Then
    '            myFontStyle = myFontStyle Or FontStyle.Underline
    '        End If

    '        Chart1.ChartAreas(0).AxisY.TitleFont = New Font("Arial", StockChart.YAxis.Title.Size, myFontStyle)
    '        Chart1.ChartAreas(0).AxisY.Title = StockChart.YAxis.Title.Text

    '        'Draw Chart Label:
    '        'Check if "Label1" is already in the list of titles:
    '        If Chart1.Titles.IndexOf("Label1") = -1 Then 'Label "Label1" doesnt exist
    '            Chart1.Titles.Add("Label1").Name = "Label1" 'The name needs to be explicitly declared!
    '        End If

    '        Chart1.Titles("Label1").Text = StockChart.ChartLabel.Text

    '        Dim myFontStyle2 As FontStyle = FontStyle.Regular
    '        If StockChart.ChartLabel.Bold Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Bold
    '        End If
    '        If StockChart.ChartLabel.Italic Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Italic
    '        End If
    '        If StockChart.ChartLabel.Strikeout Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Strikeout
    '        End If
    '        If StockChart.ChartLabel.Underline Then
    '            myFontStyle2 = myFontStyle2 Or FontStyle.Underline
    '        End If

    '        Chart1.Titles("Label1").Font = New Font("Arial", StockChart.ChartLabel.Size, myFontStyle2)
    '        Chart1.Titles("Label1").Alignment = StockChart.ChartLabel.Alignment

    '        Chart1.ChartAreas(0).AxisX.LabelStyle.IsEndLabelVisible = True

    '    Catch ex As Exception
    '        Message.AddWarning("Error drawing stock chart: " & ex.Message & vbCrLf)
    '    End Try

    'End Sub


    'Private Sub UpdateStockChartSettings()
    '    'Update StockChart with the settings selected on the Chart Type, Titles, X Axis and Y Axis tabs.

    '    'Update Chart Properties:

    '    If txtSeriesName.Text <> "" Then
    '        StockChart.SeriesName = Trim(txtSeriesName.Text)
    '    End If

    '    If txtDataDescription.Text <> "" Then
    '        StockChart.InputDataDescr = txtDataDescription.Text
    '    End If

    '    If IsNothing(cmbXValues.SelectedItem) Then
    '        Message.AddWarning("The Field containing the XValues for the chart has not been selected." & vbCrLf)
    '    Else
    '        StockChart.XValuesFieldName = cmbXValues.SelectedItem.ToString
    '    End If

    '    If Trim(DataGridView1.Rows(0).Cells(1).Value) = "" Then
    '        Message.AddWarning("The Field containing the YValues High for the chart has not been selected." & vbCrLf)
    '    Else
    '        StockChart.YValuesHighFieldName = DataGridView1.Rows(0).Cells(1).Value
    '    End If
    '    If Trim(DataGridView1.Rows(1).Cells(1).Value) = "" Then
    '        Message.AddWarning("The Field containing the YValues Low for the chart has not been selected." & vbCrLf)
    '    Else
    '        StockChart.YValuesLowFieldName = DataGridView1.Rows(1).Cells(1).Value
    '    End If
    '    If Trim(DataGridView1.Rows(2).Cells(1).Value) = "" Then
    '        Message.AddWarning("The Field containing the YValues Open for the chart has not been selected." & vbCrLf)
    '    Else
    '        StockChart.YValuesOpenFieldName = DataGridView1.Rows(2).Cells(1).Value
    '    End If
    '    If Trim(DataGridView1.Rows(3).Cells(1).Value) = "" Then
    '        Message.AddWarning("The Field containing the YValues Close for the chart has not been selected." & vbCrLf)
    '    Else
    '        StockChart.YValuesCloseFieldName = DataGridView1.Rows(3).Cells(1).Value
    '    End If
    '    If Trim(DataGridView2.Rows(0).Cells(2).Value) = "" Then 'LabelValueType not specified
    '    Else
    '        StockChart.LabelValueType = DataGridView2.Rows(0).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(1).Cells(2).Value) = "" Then 'MaxPixelPointWidth not specified
    '    Else
    '        StockChart.MaxPixelPointWidth = DataGridView2.Rows(1).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(2).Cells(2).Value) = "" Then 'MinPixelPointWidth not specified
    '    Else
    '        StockChart.MinPixelPointWidth = DataGridView2.Rows(2).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(3).Cells(2).Value) = "" Then 'OpenCloseStyle not specified
    '    Else
    '        StockChart.OpenCloseStyle = DataGridView2.Rows(3).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(4).Cells(2).Value) = "" Then 'PixelPointDepth not specified
    '    Else
    '        StockChart.PixelPointDepth = DataGridView2.Rows(4).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(5).Cells(2).Value) = "" Then 'PixelPointGapDepth not specified
    '    Else
    '        StockChart.PixelPointGapDepth = DataGridView2.Rows(5).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(6).Cells(2).Value) = "" Then 'PixelPointWidth not specified
    '    Else
    '        StockChart.PixelPointWidth = DataGridView2.Rows(6).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(7).Cells(2).Value) = "" Then 'PointWidth not specified
    '    Else
    '        StockChart.PointWidth = DataGridView2.Rows(7).Cells(2).Value
    '    End If
    '    If Trim(DataGridView2.Rows(8).Cells(2).Value) = "" Then 'ShowOpenClose not specified
    '    Else
    '        StockChart.ShowOpenClose = DataGridView2.Rows(8).Cells(2).Value
    '    End If


    '    StockChart.ChartLabel.FontName = txtChartTitle.Font.Name
    '    StockChart.ChartLabel.Size = txtChartTitle.Font.Size
    '    StockChart.ChartLabel.Bold = txtChartTitle.Font.Bold
    '    StockChart.ChartLabel.Italic = txtChartTitle.Font.Italic
    '    StockChart.ChartLabel.Strikeout = txtChartTitle.Font.Strikeout
    '    StockChart.ChartLabel.Underline = txtChartTitle.Font.Underline
    '    StockChart.ChartLabel.Text = txtChartTitle.Text

    '    StockChart.ChartLabel.Color = txtChartTitle.ForeColor

    '    If IsNothing(cmbAlignment.SelectedItem) Then
    '    Else
    '        StockChart.ChartLabel.Alignment = [Enum].Parse(GetType(ContentAlignment), cmbAlignment.SelectedItem.ToString)
    '    End If

    '    'Update X Axis settings:
    '    StockChart.XAxis.Title.FontName = txtXAxisTitle.Font.Name
    '    StockChart.XAxis.Title.Size = txtXAxisTitle.Font.Size
    '    StockChart.XAxis.Title.Bold = txtXAxisTitle.Font.Bold
    '    StockChart.XAxis.Title.Italic = txtXAxisTitle.Font.Italic
    '    StockChart.XAxis.Title.Strikeout = txtXAxisTitle.Font.Strikeout
    '    StockChart.XAxis.Title.Underline = txtXAxisTitle.Font.Underline
    '    StockChart.XAxis.Title.Text = txtXAxisTitle.Text

    '    If chkXAxisAutoMin.Checked = True Then
    '        StockChart.XAxis.AutoMinimum = True
    '    Else
    '        StockChart.XAxis.AutoMinimum = False
    '    End If

    '    If chkXAxisAutoMax.Checked = True Then
    '        StockChart.XAxis.AutoMaximum = True
    '    Else
    '        StockChart.XAxis.AutoMaximum = False
    '    End If

    '    StockChart.XAxis.Minimum = Val(txtXAxisMin.Text)
    '    StockChart.XAxis.Maximum = Val(txtXAxisMax.Text)

    '    If chkXAxisAutoAnnotInt.Checked = True Then
    '        StockChart.XAxis.Interval = 0 '0 indicates auto annotation.
    '        StockChart.XAxis.AutoInterval = True
    '    Else
    '        StockChart.XAxis.Interval = Val(txtXAxisAnnotInt.Text)
    '        StockChart.XAxis.AutoInterval = False
    '    End If

    '    If chkXAxisAutoMajGridInt.Checked = True Then
    '        StockChart.XAxis.MajorGridInterval = 0
    '        StockChart.XAxis.AutoMajorGridInterval = True
    '    Else
    '        StockChart.XAxis.MajorGridInterval = Val(txtXAxisMajGridInt.Text)
    '        StockChart.XAxis.AutoMajorGridInterval = False
    '    End If


    '    'Update Y Axis settings:
    '    StockChart.YAxis.Title.FontName = txtYAxisTitle.Font.Name
    '    StockChart.YAxis.Title.Size = txtYAxisTitle.Font.Size
    '    StockChart.YAxis.Title.Bold = txtYAxisTitle.Font.Bold
    '    StockChart.YAxis.Title.Italic = txtYAxisTitle.Font.Italic
    '    StockChart.YAxis.Title.Strikeout = txtYAxisTitle.Font.Strikeout
    '    StockChart.YAxis.Title.Underline = txtYAxisTitle.Font.Underline

    '    StockChart.YAxis.Title.Text = txtYAxisTitle.Text

    '    If chkYAxisAutoMin.Checked = True Then
    '        StockChart.YAxis.AutoMinimum = True
    '    Else
    '        StockChart.YAxis.AutoMinimum = False
    '    End If

    '    If chkYAxisAutoMax.Checked = True Then
    '        StockChart.YAxis.AutoMaximum = True
    '    Else
    '        StockChart.YAxis.AutoMaximum = False
    '    End If

    '    StockChart.YAxis.Minimum = Val(txtYAxisMin.Text)
    '    StockChart.YAxis.Maximum = Val(txtYAxisMax.Text)

    '    If chkYAxisAutoAnnotInt.Checked = True Then
    '        StockChart.YAxis.Interval = 0 '0 indicates auto annotation.
    '        StockChart.YAxis.AutoInterval = True
    '    Else
    '        StockChart.YAxis.Interval = Val(txtYAxisAnnotInt.Text)
    '        StockChart.YAxis.AutoInterval = False
    '    End If

    '    If chkYAxisAutoMajGridInt.Checked = True Then
    '        StockChart.YAxis.MajorGridInterval = 0
    '        StockChart.YAxis.AutoMajorGridInterval = True
    '        'Message.Add("Y Axis major grid interval set to auto." & vbCrLf)
    '    Else
    '        StockChart.YAxis.MajorGridInterval = Val(txtYAxisMajGridInt.Text)
    '        StockChart.YAxis.AutoMajorGridInterval = False
    '        'Message.Add("Y Axis major grid interval set to:" & txtYAxisMajGridInt.Text & vbCrLf)
    '    End If


    'End Sub

    Private Sub btnChartTitleFont_Click(sender As Object, e As EventArgs)
        'Edit chart title font
        FontDialog1.Font = txtChartTitle.Font
        FontDialog1.ShowDialog()
        txtChartTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
            If IsNothing(ProjectArchive) Then
                ProjectArchive = New frmArchive
                ProjectArchive.Show()
                ProjectArchive.Title = "Project Archive"
                ProjectArchive.Path = Project.Path
            Else
                ProjectArchive.Show()
                ProjectArchive.BringToFront()
            End If
        Else
            Process.Start(Project.Path)
        End If
    End Sub

    Private Sub ProjectArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ProjectArchive.FormClosed
        ProjectArchive = Nothing
    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        ElseIf Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
            If IsNothing(SettingsArchive) Then
                SettingsArchive = New frmArchive
                SettingsArchive.Show()
                SettingsArchive.Title = "Settings Archive"
                SettingsArchive.Path = Project.SettingsLocn.Path
            Else
                SettingsArchive.Show()
                SettingsArchive.BringToFront()
            End If
        End If
    End Sub

    Private Sub SettingsArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SettingsArchive.FormClosed
        SettingsArchive = Nothing
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        ElseIf Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
            If IsNothing(DataArchive) Then
                DataArchive = New frmArchive
                DataArchive.Show()
                DataArchive.Title = "Data Archive"
                DataArchive.Path = Project.DataLocn.Path
            Else
                DataArchive.Show()
                DataArchive.BringToFront()
            End If
        End If
    End Sub

    Private Sub DataArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles DataArchive.FormClosed
        DataArchive = Nothing
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        ElseIf Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
            If IsNothing(SystemArchive) Then
                SystemArchive = New frmArchive
                SystemArchive.Show()
                SystemArchive.Title = "System Archive"
                SystemArchive.Path = Project.SystemLocn.Path
            Else
                SystemArchive.Show()
                SystemArchive.BringToFront()
            End If
        End If
    End Sub

    Private Sub SystemArchive_FormClosed(sender As Object, e As FormClosedEventArgs) Handles SystemArchive.FormClosed
        SystemArchive = Nothing
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click
        Process.Start(ApplicationInfo.ApplicationDir)
    End Sub

    Private Sub btnShowProjectInfo_Click(sender As Object, e As EventArgs) Handles btnShowProjectInfo.Click
        'Show the current Project information:
        Message.Add("--------------------------------------------------------------------------------------" & vbCrLf)
        Message.Add("Project ------------------------ " & vbCrLf)
        Message.Add("   Name: " & Project.Name & vbCrLf)
        Message.Add("   Type: " & Project.Type.ToString & vbCrLf)
        Message.Add("   Description: " & Project.Description & vbCrLf)
        Message.Add("   Creation Date: " & Project.CreationDate & vbCrLf)
        Message.Add("   ID: " & Project.ID & vbCrLf)
        Message.Add("   Relative Path: " & Project.RelativePath & vbCrLf)
        Message.Add("   Path: " & Project.Path & vbCrLf & vbCrLf)

        Message.Add("Parent Project ----------------- " & vbCrLf)
        Message.Add("   Name: " & Project.ParentProjectName & vbCrLf)
        Message.Add("   Path: " & Project.ParentProjectPath & vbCrLf)

        Message.Add("Application -------------------- " & vbCrLf)
        Message.Add("   Name: " & Project.Application.Name & vbCrLf)
        Message.Add("   Description: " & Project.Application.Description & vbCrLf)
        Message.Add("   Path: " & Project.ApplicationDir & vbCrLf)

        Message.Add("Settings ----------------------- " & vbCrLf)
        Message.Add("   Settings Relative Location Type: " & Project.SettingsRelLocn.Type.ToString & vbCrLf)
        Message.Add("   Settings Relative Location Path: " & Project.SettingsRelLocn.Path & vbCrLf)
        Message.Add("   Settings Location Type: " & Project.SettingsLocn.Type.ToString & vbCrLf)
        Message.Add("   Settings Location Path: " & Project.SettingsLocn.Path & vbCrLf)

        Message.Add("Data --------------------------- " & vbCrLf)
        Message.Add("   Data Relative Location Type: " & Project.DataRelLocn.Type.ToString & vbCrLf)
        Message.Add("   Data Relative Location Path: " & Project.DataRelLocn.Path & vbCrLf)
        Message.Add("   Data Location Type: " & Project.DataLocn.Type.ToString & vbCrLf)
        Message.Add("   Data Location Path: " & Project.DataLocn.Path & vbCrLf)

        Message.Add("System ------------------------- " & vbCrLf)
        Message.Add("   System Relative Location Type: " & Project.SystemRelLocn.Type.ToString & vbCrLf)
        Message.Add("   System Relative Location Path: " & Project.SystemRelLocn.Path & vbCrLf)
        Message.Add("   System Location Type: " & Project.SystemLocn.Type.ToString & vbCrLf)
        Message.Add("   System Location Path: " & Project.SystemLocn.Path & vbCrLf)
        Message.Add("======================================================================================" & vbCrLf)

    End Sub

    Private Sub btnOpenParentDir_Click(sender As Object, e As EventArgs) Handles btnOpenParentDir.Click
        'Open the Parent directory of the selected project.
        Dim ParentDir As String = System.IO.Directory.GetParent(Project.Path).FullName
        If System.IO.Directory.Exists(ParentDir) Then
            Process.Start(ParentDir)
        Else
            Message.AddWarning("The parent directory was not found: " & ParentDir & vbCrLf)
        End If
    End Sub

    Private Sub btnCreateArchive_Click(sender As Object, e As EventArgs) Handles btnCreateArchive.Click
        'Create a Project Archive file.
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
            Message.Add("The Project is an Archive type. It is already in an archived format." & vbCrLf)

        Else
            'The project is contained in the directory Project.Path.
            'This directory and contents will be saved in a zip file in the parent directory with the same name but with extension .AdvlArchive.

            Dim ParentDir As String = System.IO.Directory.GetParent(Project.Path).FullName
            Dim ProjectArchiveName As String = System.IO.Path.GetFileName(Project.Path) & ".AdvlArchive"

            If My.Computer.FileSystem.FileExists(ParentDir & "\" & ProjectArchiveName) Then 'The Project Archive file already exists.
                Message.Add("The Project Archive file already exists: " & ParentDir & "\" & ProjectArchiveName & vbCrLf)
            Else 'The Project Archive file does not exist. OK to create the Archive.
                System.IO.Compression.ZipFile.CreateFromDirectory(Project.Path, ParentDir & "\" & ProjectArchiveName)

                'Remove all Lock files:
                Dim Zip As System.IO.Compression.ZipArchive
                Zip = System.IO.Compression.ZipFile.Open(ParentDir & "\" & ProjectArchiveName, IO.Compression.ZipArchiveMode.Update)
                Dim DeleteList As New List(Of String) 'List of entry names to delete
                Dim myEntry As System.IO.Compression.ZipArchiveEntry
                For Each entry As System.IO.Compression.ZipArchiveEntry In Zip.Entries
                    If entry.Name = "Project.Lock" Then
                        DeleteList.Add(entry.FullName)
                    End If
                Next
                For Each item In DeleteList
                    myEntry = Zip.GetEntry(item)
                    myEntry.Delete()
                Next
                Zip.Dispose()

                Message.Add("Project Archive file created: " & ParentDir & "\" & ProjectArchiveName & vbCrLf)
            End If
        End If
    End Sub

    Private Sub btnOpenArchive_Click(sender As Object, e As EventArgs) Handles btnOpenArchive.Click
        'Open a Project Archive file.

        'Use the OpenFileDialog to look for an .AdvlArchive file.      
        OpenFileDialog1.Title = "Select an Archived Project File"
        OpenFileDialog1.InitialDirectory = System.IO.Directory.GetParent(Project.Path).FullName 'Start looking in the ParentDir.
        OpenFileDialog1.Filter = "Archived Project|*.AdvlArchive"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim FileName As String = OpenFileDialog1.FileName
            OpenArchivedProject(FileName)
        End If
    End Sub

    Private Sub OpenArchivedProject(ByVal FilePath As String)
        'Open the archived project at the specified path.

        Dim Zip As System.IO.Compression.ZipArchive
        Try
            Zip = System.IO.Compression.ZipFile.OpenRead(FilePath)

            Dim Entry As System.IO.Compression.ZipArchiveEntry = Zip.GetEntry("Project_Info_ADVL_2.xml")
            If IsNothing(Entry) Then
                Message.AddWarning("The file is not an Archived Andorville Project." & vbCrLf)
                'Check if it is an Archive project type with a .AdvlProject extension.
                'NOTE: These are already zip files so no need to archive.

            Else
                Message.Add("The file is an Archived Andorville Project." & vbCrLf)
                Dim ParentDir As String = System.IO.Directory.GetParent(FilePath).FullName
                Dim ProjectName As String = System.IO.Path.GetFileNameWithoutExtension(FilePath)
                Message.Add("The Project will be expanded in the directory: " & ParentDir & vbCrLf)
                Message.Add("The Project name will be: " & ProjectName & vbCrLf)
                Zip.Dispose()
                If System.IO.Directory.Exists(ParentDir & "\" & ProjectName) Then
                    Message.AddWarning("The Project already exists: " & ParentDir & "\" & ProjectName & vbCrLf)
                Else
                    System.IO.Compression.ZipFile.ExtractToDirectory(FilePath, ParentDir & "\" & ProjectName) 'Extract the project from the archive                   
                    Project.AddProjectToList(ParentDir & "\" & ProjectName)
                    'Open the new project                 
                    CloseProject()  'Close the current project
                    Project.SelectProject(ParentDir & "\" & ProjectName) 'Select the project at the specifed path.
                    OpenProject() 'Open the selected project.
                End If
            End If
        Catch ex As Exception
            Message.AddWarning("Error opening Archived Andorville Project: " & ex.Message & vbCrLf)
        End Try
    End Sub


    Private Sub TabPage2_DragEnter(sender As Object, e As DragEventArgs) Handles TabPage2.DragEnter
        'DragEnter: An object has been dragged into TabPage2 - Project Information tab.
        'This code is required to get the link to the item(s) being dragged into Project Information:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If
    End Sub

    Private Sub TabPage2_DragDrop(sender As Object, e As DragEventArgs) Handles TabPage2.DragDrop
        'A file has been dropped into the Project Information tab.

        Dim Path As String()
        Path = e.Data.GetData(DataFormats.FileDrop)
        Dim I As Integer

        If Path.Count > 0 Then
            If Path.Count > 1 Then
                Message.AddWarning("More than one file has been dropped into the Project Information tab. Only the first one will be opened." & vbCrLf)
            End If

            Try
                Dim ArchivedProjectPath As String = Path(0)
                If ArchivedProjectPath.EndsWith(".AdvlArchive") Then
                    Message.Add("The archived project will be opened: " & vbCrLf & ArchivedProjectPath & vbCrLf)
                    OpenArchivedProject(ArchivedProjectPath)
                Else
                    Message.Add("The dropped file is not an archived project: " & vbCrLf & ArchivedProjectPath & vbCrLf)
                End If
            Catch ex As Exception
                Message.AddWarning("Error opening dropped archived project. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Save the Stock Chart file.

        'Dim ChartFileName As String = Trim(txtChartFileName.Text)
        'StockChart.SaveFile(ChartFileName)

        Dim FileName As String = Trim(txtChartFileName.Text)

        If FileName = "" Then
            Message.AddWarning("Please enter a file name." & vbCrLf)
            Exit Sub
        End If

        If LCase(FileName).EndsWith(".stockchart") Then
            FileName = IO.Path.GetFileNameWithoutExtension(FileName) & ".StockChart"
        ElseIf FileName.Contains(".") Then
            Message.AddWarning("Unknown file extension: " & IO.Path.GetExtension(FileName) & vbCrLf)
            Exit Sub
        Else
            FileName = FileName & ".StockChart"
        End If

        txtChartFileName.Text = FileName
        Project.SaveXmlData(FileName, ChartInfo.ToXDoc(Chart1))

    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        'Open Stock chart file.

        'Find and open a Stock Chart file:
        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select a Stock Chart file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "Line Chart files | *.StockChart"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim FileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    txtChartFileName.Text = FileName
                    Try
                        ChartInfo.LoadFile(FileName, Chart1)
                    Catch ex As Exception
                        Message.AddWarning("Error in ChartInfo.LoadFile. FileName = " & FileName & vbCrLf)
                        Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
                    End Try

                    UpdateInputDataTabSettings()
                    UpdateTitlesTabSettings()
                    UpdateAreasTabSettings() 'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                    UpdateSeriesTabSettings()
                    If chkAutoDraw.Checked Then DrawStockChart()
                End If

            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select a Stock Chart file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile() 'Show the Select File form
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtensions = {".StockChart"}
                Zip.SelectFileForm.GetFileList()
                If Zip.SelectedFile <> "" Then
                    'A file has been selected
                    txtChartFileName.Text = Zip.SelectedFile
                    Try
                        ChartInfo.LoadFile(Zip.SelectedFile, Chart1)
                    Catch ex As Exception
                        Message.AddWarning("Error in ChartInfo.LoadFile. Zip.SelectedFile = " & Zip.SelectedFile & vbCrLf)
                        Message.AddWarning(ex.Message & vbCrLf & vbCrLf)
                    End Try
                    UpdateInputDataTabSettings()
                    UpdateTitlesTabSettings()
                    UpdateAreasTabSettings()  'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
                    UpdateSeriesTabSettings()
                    If chkAutoDraw.Checked Then DrawStockChart()
                End If
        End Select


        ''Find and open a chart file.
        'Select Case Project.DataLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        'Select a Stock Chart file from the project directory:
        '        OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
        '        OpenFileDialog1.Filter = "Stock Chart files | *.StockChart"
        '        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
        '            Dim FileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
        '            txtChartFileName.Text = FileName

        '            StockChart.LoadFile(FileName)
        '            StockChart.FileName = FileName
        '            UpdateStockChartForm()
        '            If chkAutoDraw.Checked Then DrawStockChart()
        '        End If
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        'Select a Point Chart file from the project archive:
        '        'Show the zip archive file selection form:
        '        Zip = New ADVL_Utilities_Library_1.ZipComp
        '        Zip.ArchivePath = Project.DataLocn.Path
        '        Zip.SelectFile() 'Show the Select File form
        '        'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
        '        Zip.SelectFileForm.ApplicationName = Project.Application.Name
        '        Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
        '        Zip.SelectFileForm.Show()
        '        Zip.SelectFileForm.RestoreFormSettings()
        '        Zip.SelectFileForm.FileExtensions = {".StockChart"}
        '        Zip.SelectFileForm.GetFileList()
        '        If Zip.SelectedFile <> "" Then
        '            'A file has been selected
        '            txtChartFileName.Text = Zip.SelectedFile

        '            'OpenChart(Zip.SelectedFile)
        '            StockChart.LoadFile(Zip.SelectedFile)
        '            StockChart.FileName = Zip.SelectedFile
        '            UpdateStockChartForm()
        '            If chkAutoDraw.Checked Then DrawStockChart()
        '        End If
        'End Select
    End Sub



    Private Sub txtChartFileName_LostFocus(sender As Object, e As EventArgs) Handles txtChartFileName.LostFocus
        Dim FileName As String = Trim(txtChartFileName.Text)
        If FileName.EndsWith(".StockChart") Then
            If FileName = ".StockChart" Then
                Message.AddWarning("The chart file name is blank!" & vbCrLf)
            Else
                'StockChart.FileName = FileName
                ChartInfo.FileName = FileName
                txtChartFileName.Text = FileName
            End If
        Else
            If FileName.Contains(".") Then
                'Remove the wrong file extension.
                Dim Pos As Integer = InStr(FileName, ".")
                FileName = Microsoft.VisualBasic.Left(FileName, Pos - 1)
                If FileName = "" Then
                    Message.AddWarning("The chart file name is blank!" & vbCrLf)
                Else
                    FileName = FileName & ".StockChart"
                    'StockChart.FileName = FileName
                    ChartInfo.FileName = FileName
                    txtChartFileName.Text = FileName
                End If
            Else
                If FileName = "" Then
                    Message.AddWarning("The chart file name is blank!" & vbCrLf)
                Else
                    FileName = FileName & ".StockChart"
                    'StockChart.FileName = FileName
                    ChartInfo.FileName = FileName
                    txtChartFileName.Text = FileName
                End If
            End If
        End If

    End Sub



    Private Sub chkConnect_LostFocus(sender As Object, e As EventArgs) Handles chkConnect.LostFocus
        If chkConnect.Checked Then
            Project.ConnectOnOpen = True
        Else
            Project.ConnectOnOpen = False
        End If
        Project.SaveProjectInfoFile()
    End Sub


#End Region 'Process XMessages ------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub ToolStripMenuItem1_EditWorkflowTabPage_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_EditWorkflowTabPage.Click
        'Edit the Workflow Web Page:

        If WorkflowFileName = "" Then
            Message.AddWarning("No page to edit." & vbCrLf)
        Else
            Dim FormNo As Integer = OpenNewHtmlDisplayPage()
            HtmlDisplayFormList(FormNo).FileName = WorkflowFileName
            HtmlDisplayFormList(FormNo).OpenDocument
        End If

    End Sub

    Private Sub ToolStripMenuItem1_ShowStartPageInWorkflowTab_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_ShowStartPageInWorkflowTab.Click
        'Show the Start Page in the Workflow Tab:
        OpenStartPage()

    End Sub

    Private Sub bgwComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwComCheck.DoWork
        'The communications check thread.

        While ConnectedToComNet
            Try
                If client.IsAlive() Then
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                Else
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection Fault.")
                End If
            Catch ex As Exception
                bgwComCheck.ReportProgress(1, "Error in bgeComCheck_DoWork!" & vbCrLf)
                bgwComCheck.ReportProgress(1, ex.Message & vbCrLf)
            End Try

            'System.Threading.Thread.Sleep(60000) 'Sleep time in milliseconds (60 seconds) - For testing only.
            'System.Threading.Thread.Sleep(3600000) 'Sleep time in milliseconds (60 minutes)
            System.Threading.Thread.Sleep(1800000) 'Sleep time in milliseconds (30 minutes)
        End While
    End Sub

    Private Sub bgwComCheck_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwComCheck.ProgressChanged
        Message.Add(e.UserState.ToString) 'Show the ComCheck message 
    End Sub

    Private Sub bgwSendMessage_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessage.DoWork
        'Send a message on a separate thread:
        Try
            If IsNothing(client) Then
                bgwSendMessage.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessage.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParams As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParams.ProjectNetworkName, SendMessageParams.ConnectionName, SendMessageParams.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessage.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessage_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessage.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub bgwSendMessageAlt_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessageAlt.DoWork
        'Alternative SendMessage background worker - to send a message while instructions are being processed. 
        'Send a message on a separate thread
        Try
            If IsNothing(client) Then
                bgwSendMessageAlt.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessageAlt.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParamsAlt As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParamsAlt.ProjectNetworkName, SendMessageParamsAlt.ConnectionName, SendMessageParamsAlt.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessageAlt.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessageAlt_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessageAlt.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub XMsg_ErrorMsg(ErrMsg As String) Handles XMsg.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XMsgLocal_Instruction(Info As String, Locn As String) Handles XMsgLocal.Instruction

    End Sub

    Private Sub Message_ShowXMessagesChanged(Show As Boolean) Handles Message.ShowXMessagesChanged
        ShowXMessages = Show
    End Sub

    Private Sub Message_ShowSysMessagesChanged(Show As Boolean) Handles Message.ShowSysMessagesChanged
        ShowSysMessages = Show
    End Sub

#Region " Chart Input Data Tab" '========================================================================================================================================

    Private Sub UpdateInputDataTabSettings()
        'Update the Input Data tab settings from the Chart control.

        txtDatabasePath.Text = ChartInfo.InputDatabasePath
        Select Case ChartInfo.InputDataType
            Case "Database"
                rbDatabase.Checked = True
            Case "Dataset"
                rbDataset.Checked = True
            Case Else
                rbDatabase.Checked = True
        End Select
        FillLstTables()
        txtDataDescription.Text = ChartInfo.InputDataDescr
        txtInputQuery.Text = ChartInfo.InputQuery

        ChartInfo.ApplyQuery()
        'GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
        UpdateSeriesTabSettings()
        GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.

    End Sub



    Private Sub btnDatabase_Click(sender As Object, e As EventArgs) Handles btnDatabase.Click
        'Select a database file:

        OpenFileDialog1.Filter = "Access Database |*.accdb"
        OpenFileDialog1.FileName = ""

        If InputDatabaseDirectory <> "" Then
            OpenFileDialog1.InitialDirectory = InputDatabaseDirectory
        End If

        OpenFileDialog1.ShowDialog()

        'PointChart.InputDatabasePath = OpenFileDialog1.FileName
        ChartInfo.InputDatabasePath = OpenFileDialog1.FileName
        'txtDatabasePath.Text = PointChart.InputDatabasePath
        txtDatabasePath.Text = ChartInfo.InputDatabasePath
        'FillLstTables()

        InputDatabaseDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)

        Message.Add("InputDatabaseDirectory = " & InputDatabaseDirectory & vbCrLf)

        FillLstTables()
    End Sub

    Private Sub FillLstTables()
        'Fill the lstSelectTable listbox with the availalble tables in the selected database.

        lstTables.Items.Clear()

        'If PointChart.InputDatabasePath = "" Then Exit Sub
        If ChartInfo.InputDatabasePath = "" Then Exit Sub

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim dt As DataTable

        lstTables.Items.Clear()

        'Specify the connection string:
        'Access 2003
        'connectionString = "provider=Microsoft.Jet.OLEDB.4.0;" + _
        '"data source = " + txtDatabase.Text

        'Access 2007:
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + txtDatabasePath.Text

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)

        conn.Open()

        Dim restrictions As String() = New String() {Nothing, Nothing, Nothing, "TABLE"} 'This restriction removes system tables
        dt = conn.GetSchema("Tables", restrictions)

        'Fill lstTables
        Dim dr As DataRow
        Dim I As Integer 'Loop index
        Dim MaxI As Integer

        MaxI = dt.Rows.Count
        For I = 0 To MaxI - 1
            dr = dt.Rows(0)
            lstTables.Items.Add(dt.Rows(I).Item(2).ToString)
        Next I

        conn.Close()

    End Sub

    Private Sub lstTables_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTables.SelectedIndexChanged
        FillLstFields()
    End Sub

    Private Sub FillLstFields()
        'Fill the lstFields listbox with the availalble fields in the selected table.

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.
        Dim ds As DataSet 'Declate a Dataset.
        Dim dt As DataTable

        If lstTables.SelectedIndex = -1 Then 'No item is selected
            lstFields.Items.Clear()

        Else 'A table has been selected. List its fields:
            lstFields.Items.Clear()

            'Specify the connection string (Access 2007):
            connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
            "data source = " + txtDatabasePath.Text

            'Connect to the Access database:
            conn = New System.Data.OleDb.OleDbConnection(connectionString)
            conn.Open()

            'Specify the commandString to query the database:
            commandString = "SELECT TOP 500 * FROM " + lstTables.SelectedItem.ToString
            Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)
            ds = New DataSet
            dataAdapter.Fill(ds, "SelTable") 'ds was defined earlier as a DataSet
            dt = ds.Tables("SelTable")

            Dim NFields As Integer
            NFields = dt.Columns.Count
            Dim I As Integer
            For I = 0 To NFields - 1
                lstFields.Items.Add(dt.Columns(I).ColumnName.ToString)
            Next

            conn.Close()

        End If
    End Sub

    Private Sub btnViewData_Click(sender As Object, e As EventArgs) Handles btnViewData.Click
        'Open the View Database Data form:
        If IsNothing(ViewDatabaseData) Then
            ViewDatabaseData = New frmViewDatabaseData
            ViewDatabaseData.Show()
            ViewDatabaseData.UpdateTable()
        Else
            ViewDatabaseData.Show()
            ViewDatabaseData.UpdateTable()
        End If
    End Sub

    Private Sub ViewDatabaseData_FormClosed(sender As Object, e As FormClosedEventArgs) Handles ViewDatabaseData.FormClosed
        ViewDatabaseData = Nothing
    End Sub

    Private Sub btnApplyQuery_Click(sender As Object, e As EventArgs) Handles btnApplyQuery.Click
        'PointChart.InputQuery = txtInputQuery.Text
        ChartInfo.InputQuery = txtInputQuery.Text
        ApplyQuery() 'This fills the Dataset with the Query data.

        'THEFOLLOWING CODE is now included in ApplyQuery():
        'GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
        'UpdateSeriesTabSettings()
    End Sub

    Private Sub GetChartOptionsFromDataset()
        'Update the Chart display options from the Dataset.

        'First check that data to chart has been loaded into the dataset:
        'If ds.Tables.Count = 0 Then
        If ChartInfo.ds.Tables.Count = 0 Then
            Message.AddWarning("No data has been selected for charting." & vbCrLf)
        Else
            'Get the list of available fields from the dataset:
            GetFieldListFromDataset()
        End If

        'Show the selected XValues field: --------------------------------------------------------------------------------
        'Dim I As Integer 'Loop index
        'For I = 1 To cmbXValues.Items.Count
        '    If cmbXValues.Items(I - 1) = PointChart.XValuesFieldName Then
        '        cmbXValues.SelectedIndex = I - 1
        '    End If
        'Next

        Dim SeriesName As String = txtSeriesName.Text.Trim
        If SeriesName = "" Then
            'Message.AddWarning("The Series Name is blank" & vbCrLf)
            'ElseIf ChartInfo.dictFields.ContainsKey(SeriesName) Then
        ElseIf ChartInfo.dictSeriesInfo.ContainsKey(SeriesName) Then
            Dim I As Integer 'Loop index
            For I = 1 To cmbXValues.Items.Count
                'The Series Name is in txtSeriesName.Text
                If cmbXValues.Items(I - 1) = ChartInfo.dictSeriesInfo(txtSeriesName.Text).XValuesFieldName Then
                    cmbXValues.SelectedIndex = I - 1
                End If
            Next
        Else
            Message.AddWarning("The Series Name is not in the Chart Info dictionary: " & SeriesName & vbCrLf)
        End If

    End Sub

    Private Sub GetFieldListFromDataset()
        'Update the available list of fields for plotting on the X and Y axes.

        'cboFieldSelections.Items.Clear()
        cmbXValues.Items.Clear()
        cboFieldSelections.Items.Clear()

        'If ds.Tables(0).Columns.Count > 0 Then
        If ChartInfo.ds.Tables(0).Columns.Count > 0 Then
            Dim I As Integer 'Loop index
            'For I = 1 To ds.Tables(0).Columns.Count
            For I = 1 To ChartInfo.ds.Tables(0).Columns.Count
                cboFieldSelections.Items.Add(ChartInfo.ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox column used in DataGridView1 (Y Values to chart)
                cmbXValues.Items.Add(ChartInfo.ds.Tables(0).Columns(I - 1).ColumnName) 'ComboBox used to select the Field to use along the X Axis
            Next
        End If
    End Sub

    Public Sub ApplyQuery()
        'Use the query to fill the ds dataset

        'If PointChart.InputDatabasePath = "" Then
        If ChartInfo.InputDatabasePath = "" Then
            Message.AddWarning("InputDatabasePath is not defined!" & vbCrLf)
            Exit Sub
        End If

        'Database access for MS Access:
        Dim connectionString As String 'Declare a connection string for MS Access - defines the database or server to be used.
        Dim conn As System.Data.OleDb.OleDbConnection 'Declare a connection for MS Access - used by the Data Adapter to connect to and disconnect from the database.
        Dim commandString As String 'Declare a command string - contains the query to be passed to the database.

        'Specify the connection string (Access 2007):
        connectionString = "provider=Microsoft.ACE.OLEDB.12.0;" +
        "data source = " + ChartInfo.InputDatabasePath

        'Connect to the Access database:
        conn = New System.Data.OleDb.OleDbConnection(connectionString)
        conn.Open()

        'Specify the commandString to query the database:
        commandString = ChartInfo.InputQuery
        Dim dataAdapter As New System.Data.OleDb.OleDbDataAdapter(commandString, conn)

        ChartInfo.ds.Clear()
        ChartInfo.ds.Reset()

        dataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Try
            dataAdapter.Fill(ChartInfo.ds, "SelTable")
            'UpdateChartQuery() 'NOT NEEDED??? 'This was originally used to set PointChart or StockChart .Input Query to the property InputQuery. (See the Chart app code.)
            GetChartOptionsFromDataset() 'This updates the chart options using Dataset fields.
            UpdateSeriesTabSettings()
        Catch ex As Exception
            Message.AddWarning("Error applying query." & vbCrLf)
            Message.AddWarning(ex.Message & vbCrLf)
        End Try

        conn.Close()

    End Sub

    Private Sub txtDataDescription_LostFocus(sender As Object, e As EventArgs) Handles txtDataDescription.LostFocus
        ChartInfo.InputDataDescr = txtDataDescription.Text
    End Sub

    Private Sub txtInputQuery_LostFocus(sender As Object, e As EventArgs) Handles txtInputQuery.LostFocus
        ChartInfo.InputQuery = txtInputQuery.Text
    End Sub

#End Region 'Chart Input Data Tab ---------------------------------------------------------------------------------------------------------------------------------------


#Region " Chart Titles Tab" '============================================================================================================================================

    Private Sub UpdateTitlesTabSettings()
        'Update the Titles tab settings from the Chart control.

        Dim NTitlesRecords As Integer = Chart1.Titles.Count
        txtNTitlesRecords.Text = NTitlesRecords

        Dim TitleNo As Integer = Val(txtTitlesRecordNo.Text)

        If TitleNo + 1 > NTitlesRecords Then TitleNo = NTitlesRecords - 1

        If TitleNo >= 0 Then
            txtTitlesRecordNo.Text = TitleNo + 1
            txtChartTitle.Text = Chart1.Titles(TitleNo).Text
            txtTitleName.Text = Chart1.Titles(TitleNo).Name

            cmbAlignment.SelectedIndex = cmbAlignment.FindStringExact(Chart1.Titles(TitleNo).Alignment.ToString)
            cmbOrientation.SelectedIndex = cmbOrientation.FindStringExact(Chart1.Titles(TitleNo).TextOrientation.ToString)

            txtChartTitle.Font = Chart1.Titles(TitleNo).Font
            txtChartTitle.ForeColor = Chart1.Titles(TitleNo).ForeColor
        Else
            txtTitlesRecordNo.Text = 0
            txtChartTitle.Text = ""
            txtTitleName.Text = ""
        End If
    End Sub

    Private Sub ApplyTitlesTabSettings()
        'Apply the Titles tab settings to the Chart control.

        Dim TitleNo As Integer = Val(txtTitlesRecordNo.Text) - 1

        Dim TitleCount As Integer = Chart1.Titles.Count

        If TitleNo > TitleCount Then
            'Dim I As Integer
            'For I = 0 To TitleNo - TitleCount
            '    Chart1.Titles.Add("Label" & I)
            'Next
            Message.AddWarning("The title number is larger than to number of titles in the chart!" & vbCrLf)
            Exit Sub
        End If

        Chart1.Titles(TitleNo).Name = txtTitleName.Text.Trim
        Chart1.Titles(TitleNo).Text = txtChartTitle.Text
        Chart1.Titles(TitleNo).ForeColor = txtChartTitle.ForeColor
        Chart1.Titles(TitleNo).Font = txtChartTitle.Font

        Select Case cmbAlignment.SelectedItem.ToString
            Case "BottomCenter"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomCenter
            Case "BottomLeft"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomLeft
            Case "BottomRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomRight
            Case "MiddleCenter"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.MiddleCenter
            Case "MiddleLeft"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.MiddleLeft
            Case "MiddleRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.MiddleRight
            Case "TopCenter"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopCenter
            Case "TopLeft"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopLeft
            Case "TopRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopRight
            Case "BottomRight"
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.BottomRight
            Case Else
                Chart1.Titles(TitleNo).Alignment = ContentAlignment.TopCenter
        End Select

        Select Case cmbOrientation.SelectedItem.ToString
            Case "Auto"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Auto
            Case "Horizontal"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Horizontal
            Case "Rotated270"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Rotated270
            Case "Rotated90"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Rotated90
            Case "Stacked"
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Stacked
            Case Else
                Chart1.Titles(TitleNo).TextOrientation = DataVisualization.Charting.TextOrientation.Auto
        End Select

    End Sub

    Private Sub btnAddTitle_Click(sender As Object, e As EventArgs) Handles btnAddTitle.Click
        'Add a new Chart Title:

        Dim NewTitleNo As Integer = Chart1.Titles.Count + 1
        If NewTitleNo = 0 Then NewTitleNo = 1 'The first title should be named Title1

        Message.Add("Chart1.Titles.Count = " & Chart1.Titles.Count & vbCrLf)
        Dim NewTitleName As String = "Label" & NewTitleNo
        Chart1.Titles.Add(NewTitleName)

        txtTitleName.Text = NewTitleName
        'txtTitlesRecordNo.Text = NewTitleNo + 1
        txtTitlesRecordNo.Text = NewTitleNo
        txtNTitlesRecords.Text = Chart1.Titles.Count
        txtChartTitle.Text = ""

    End Sub

    Private Sub btnApplyTitlesSettings_Click(sender As Object, e As EventArgs) Handles btnApplyTitlesSettings.Click
        'Apply the Titles settings:
        ApplyTitlesTabSettings()
    End Sub


#End Region 'Chart Titles Tab -------------------------------------------------------------------------------------------------------------------------------------------


#Region " Chart Series Tab" '============================================================================================================================================

    'Private Sub SetupLineChartSeriesTab()
    Private Sub SetupStockChartSeriesTab()
        'Set up the series tab for a Stock Chart:

        'List of Chart Types:
        'https://msdn.microsoft.com/en-us/data/dd489233(v=vs.95)
        'Point Chart Characteristics:
        'https://msdn.microsoft.com/en-us/data/dd456684(v=vs.95)
        'Stock Chart Characteristics:
        'https://msdn.microsoft.com/en-us/data/dd456733(v=vs.95)

        'Set up Stock chart:
        txtChartDescr.Text = "Stock Chart" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "A Stock chart is typically used to illustrate significant stock price points including a stock's open, close, high, and low price points. However, this type of chart" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "can also be used to analyze scientific data, because each series of data displays high, low, open, and close values, which are typically lines or triangles. The" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "opening values are shown by the markers on the left, and the closing values are shown by the markers on the right." & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "The open and close markers can be specified using the ShowOpenClose custom attribute." & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "Custom Attributes:" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "LabelValueType - Specifies the Y value to use as the data point label. (High, Low, Open, Close)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "MaxPixelPointWidth - Specifies the maximum width of the data point in pixels. (Any integer > 0)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "MinPixelPointWidth - Specifies the minimum width of the data point in pixels. (Any integer > 0)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "OpenCloseStyle - Specifies the marker style for open and close values. (Triangle, Line, Candlestick)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "PixelPointDepth - Specifies the 3D series depth in pixels. (Any integer > 0)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "PixelPointGapDepth - Specifies the 3D gap depth in pixels. (Any integer > 0)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "PixelPointWidth - Specifies the data point width in pixels. (0 - 2)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "PointWidth - Data point width. (Any integer > 0)" & vbCrLf
        txtChartDescr.Text = txtChartDescr.Text & "ShowOpenClose - Specifies whether markers for open and close prices are displayed. (Both, Open, Close)" & vbCrLf

        'txtChartDescr.Text = txtChartDescr.Text & "ShowMarkerLines - Specifies whether marker lines are displayed when rendered in 3D. (True, False)" & vbCrLf

        'Y Values:
        'DataGridView1.Rows.Clear()
        'DataGridView1.Rows.Add(1)
        'DataGridView1.Rows(0).Cells(0).Value = "Yvalue" 'Y Value Parameter Name

        'Set up the Y Values grid:
        DataGridView1.ColumnCount = 1
        DataGridView1.RowCount = 1
        DataGridView1.Columns(0).HeaderText = "Y Value"
        DataGridView1.Columns(0).Width = 60
        DataGridView1.Columns.Insert(1, cboFieldSelections) 'cboFieldSelections is declared glabally in the Variable Declarations section.
        DataGridView1.Columns(1).HeaderText = "Field"
        DataGridView1.Columns(1).Width = 120
        DataGridView1.AllowUserToResizeColumns = True

        'Y Values:
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(4)
        DataGridView1.Rows(0).Cells(0).Value = "High" 'First Y Value Parameter Name
        DataGridView1.Rows(1).Cells(0).Value = "Low" 'Second Y Value parameter name
        DataGridView1.Rows(2).Cells(0).Value = "Open" 'Third Y Value parameter name
        DataGridView1.Rows(3).Cells(0).Value = "Close" 'Fourth Y value parameter name


        'Custom Attributes:
        'DataGridView2.Rows.Clear()
        'DataGridView2.Rows.Add(5)

        'Set up the Custom Attributes grid:
        DataGridView2.ColumnCount = 4
        DataGridView2.RowCount = 1
        DataGridView2.Columns(0).HeaderText = "Custom Attribute"
        DataGridView2.Columns(0).Width = 120
        DataGridView2.Columns(1).HeaderText = "Value Range"
        DataGridView2.Columns(1).Width = 120
        DataGridView2.Columns(2).HeaderText = "Value"
        DataGridView2.Columns(2).Width = 120
        DataGridView2.Columns(3).HeaderText = "Description"
        DataGridView2.Columns(3).Width = 340
        DataGridView2.AllowUserToResizeColumns = True





        ''  EmptyPointValue:
        'DataGridView2.Rows(0).Cells(0).Value = "EmptyPointValue"
        'DataGridView2.Rows(0).Cells(1).Value = "Average, Zero"
        'Dim cbc0 As New DataGridViewComboBoxCell
        'cbc0.Items.Add(" ")
        'cbc0.Items.Add("Average")
        'cbc0.Items.Add("Zero")
        'DataGridView2.Rows(0).Cells(2) = cbc0

        ''  LabelStyle:
        'DataGridView2.Rows(1).Cells(0).Value = "LabelStyle"
        'DataGridView2.Rows(1).Cells(1).Value = "Auto, Top, Bottom, Right, Left, TopLeft, TopRight, BottomLeft, BottomRight, Center"
        'Dim cbc1 As New DataGridViewComboBoxCell
        'cbc1.Items.Add(" ")
        'cbc1.Items.Add("Auto")
        'cbc1.Items.Add("Top")
        'cbc1.Items.Add("Bottom")
        'cbc1.Items.Add("Right")
        'cbc1.Items.Add("Left")
        'cbc1.Items.Add("TopLeft")
        'cbc1.Items.Add("TopRight")
        'cbc1.Items.Add("BottomLeft")
        'cbc1.Items.Add("BottomRight")
        'cbc1.Items.Add("Center")
        'DataGridView2.Rows(1).Cells(2) = cbc1

        ''  PixelPointDepth:
        'DataGridView2.Rows(2).Cells(0).Value = "PixelPointDepth"
        'DataGridView2.Rows(2).Cells(1).Value = "Any integer > 0"

        ''  PixelPointGapDepth:
        'DataGridView2.Rows(3).Cells(0).Value = "PixelPointGapDepth"
        'DataGridView2.Rows(3).Cells(1).Value = "Any integer > 0"

        'Custom Attributes:
        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(9)
        DataGridView2.Rows(0).Cells(0).Value = "LabelValueType"
        DataGridView2.Rows(0).Cells(1).Value = "High, Low, Open, Close"
        Dim cbc0 As New DataGridViewComboBoxCell
        cbc0.Items.Add(" ")
        cbc0.Items.Add("High")
        cbc0.Items.Add("Low")
        cbc0.Items.Add("Open")
        cbc0.Items.Add("Close")
        DataGridView2.Rows(0).Cells(2) = cbc0
        DataGridView2.Rows(0).Cells(3).Value = "Specifies the Y value to use as the data point label."
        DataGridView2.Rows(1).Cells(0).Value = "MaxPixelPointWidth"
        DataGridView2.Rows(1).Cells(1).Value = "Any integer > 0"
        DataGridView2.Rows(1).Cells(3).Value = "Specifies the maximum width of the data point in pixels."
        DataGridView2.Rows(2).Cells(0).Value = "MinPixelPointWidth"
        DataGridView2.Rows(2).Cells(1).Value = "Any integer > 0"
        DataGridView2.Rows(2).Cells(3).Value = "Specifies the minimum data point width in pixels."
        DataGridView2.Rows(3).Cells(0).Value = "OpenCloseStyle"
        DataGridView2.Rows(3).Cells(1).Value = "Triangle, Line, Candlestick"
        Dim cbc3 As New DataGridViewComboBoxCell
        cbc3.Items.Add(" ")
        cbc3.Items.Add("Triangle")
        cbc3.Items.Add("Line")
        cbc3.Items.Add("Candlestick")
        DataGridView2.Rows(3).Cells(2) = cbc3
        DataGridView2.Rows(3).Cells(3).Value = "Specifies the marker style for open and close values."
        DataGridView2.Rows(4).Cells(0).Value = "PixelPointDepth"
        DataGridView2.Rows(4).Cells(1).Value = "Any integer > 0"
        DataGridView2.Rows(4).Cells(3).Value = "Specifies the 3D series depth in pixels."
        DataGridView2.Rows(5).Cells(0).Value = "PixelPointGapDepth"
        DataGridView2.Rows(5).Cells(1).Value = "Any integer > 0"
        DataGridView2.Rows(5).Cells(3).Value = "Specifies the 3D gap depth in pixels."
        DataGridView2.Rows(6).Cells(0).Value = "PixelPointWidth"
        DataGridView2.Rows(6).Cells(1).Value = "Any integer > 0"
        DataGridView2.Rows(6).Cells(3).Value = "Specifies the data point width in pixels."
        DataGridView2.Rows(7).Cells(0).Value = "PointWidth"
        DataGridView2.Rows(7).Cells(1).Value = "0 to 2"
        DataGridView2.Rows(7).Cells(3).Value = "Specifies data point width."
        DataGridView2.Rows(8).Cells(0).Value = "ShowOpenClose"
        DataGridView2.Rows(8).Cells(1).Value = "Both, Open, Close"
        DataGridView2.Rows(8).Cells(3).Value = "Specifies whether markers for open and close prices are displayed."
        Dim cbc8 As New DataGridViewComboBoxCell
        cbc8.Items.Add(" ")
        cbc8.Items.Add("Both")
        cbc8.Items.Add("Open")
        cbc8.Items.Add("Close")
        DataGridView2.Rows(8).Cells(2) = cbc8



    End Sub

    Private Sub btnApplySeriesSettings_Click(sender As Object, e As EventArgs) Handles btnApplySeriesSettings.Click
        ApplySeriesTabSettings()
    End Sub

    Private Sub ApplySeriesTabSettings()
        'Update the Chart control settings from the Series tab.

        Dim SeriesNo As Integer = Val(txtSeriesRecordNo.Text) - 1
        Dim SeriesCount As Integer = Chart1.Series.Count

        If SeriesNo > SeriesCount Then
            Message.AddWarning("The series number is larger than to number of series in the chart!" & vbCrLf)
            Exit Sub
        End If

        Dim SeriesName As String = txtSeriesName.Text.Trim
        Chart1.Series(SeriesNo).Name = SeriesName
        If ChartInfo.dictSeriesInfo.ContainsKey(SeriesName) Then
            'SeriesName is already in the dictionary of database fields.
        Else
            ChartInfo.dictSeriesInfo.Add(SeriesName, New SeriesInfo)
        End If

        'Chart1.Series(SeriesNo).ChartType = DataVisualization.Charting.SeriesChartType.Line
        'Chart1.Series(SeriesNo).ChartType = DataVisualization.Charting.SeriesChartType.Point
        Chart1.Series(SeriesNo).ChartType = DataVisualization.Charting.SeriesChartType.Stock
        'Chart1.Series(SeriesNo).Points.DataBindXY(ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictFields(SeriesName).XValuesFieldName, ChartInfo.ds.Tables(0).DefaultView, ChartInfo.dictFields(SeriesName).YValuesFieldName)

        'Chart1.Series(SeriesNo).ChartArea = cmbChartArea.SelectedItem.ToString
        If IsNothing(cmbChartArea.SelectedItem) Then
        Else
            ChartInfo.dictSeriesInfo(SeriesName).ChartArea = cmbChartArea.SelectedItem.ToString
            Chart1.Series(SeriesNo).ChartArea = cmbChartArea.SelectedItem.ToString
        End If
        'ChartInfo.dictSeriesInfo(SeriesName).ChartArea = cmbChartArea.SelectedItem.ToString

        'ChartInfo.dictFields(SeriesName).XValuesFieldName = cmbXValues.SelectedText
        'Message.Add("cmbXValues.SelectedText" & cmbXValues.SelectedText & vbCrLf)
        'ChartInfo.dictFields(SeriesName).XValuesFieldName = cmbXValues.SelectedItem.ToString
        ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName = cmbXValues.SelectedItem.ToString
        'Message.Add("cmbXValues.SelectedItem.ToString = " & cmbXValues.SelectedItem.ToString & vbCrLf)

        'ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName = DataGridView1.Rows(0).Cells(1).Value
        ChartInfo.dictSeriesInfo(SeriesName).YValuesHighFieldName = DataGridView1.Rows(0).Cells(1).Value
        ChartInfo.dictSeriesInfo(SeriesName).YValuesLowFieldName = DataGridView1.Rows(1).Cells(1).Value
        ChartInfo.dictSeriesInfo(SeriesName).YValuesOpenFieldName = DataGridView1.Rows(2).Cells(1).Value
        ChartInfo.dictSeriesInfo(SeriesName).YValuesCloseFieldName = DataGridView1.Rows(3).Cells(1).Value

        'Select Case cmbXAxisType.SelectedText
        Select Case cmbXAxisType.SelectedItem.ToString
            Case "Primary"
                Chart1.Series(SeriesName).XAxisType = DataVisualization.Charting.AxisType.Primary
            Case "Secondary"
                Chart1.Series(SeriesName).XAxisType = DataVisualization.Charting.AxisType.Secondary
            Case Else
                Chart1.Series(SeriesName).XAxisType = DataVisualization.Charting.AxisType.Primary
        End Select

        Select Case cmbXAxisValueType.SelectedItem.ToString
            Case "Auto"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Auto
            Case "Date"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Date
            Case "DateTime"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.DateTime
            Case "DateTimeOffset"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.DateTimeOffset
            Case "Double"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Double
            Case "Int32"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Int32
            Case "Int64"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Int64
            Case "Single"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Single
            Case "String"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.String
            Case "Time"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Time
            Case "UInt32"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.UInt32
            Case "UInt64"
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.UInt64
            Case Else
                Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Auto
        End Select

        Select Case cmbYAxisType.SelectedItem.ToString
            Case "Primary"
                Chart1.Series(SeriesName).YAxisType = DataVisualization.Charting.AxisType.Primary
            Case "Secondary"
                Chart1.Series(SeriesName).YAxisType = DataVisualization.Charting.AxisType.Secondary
            Case Else
                Chart1.Series(SeriesName).YAxisType = DataVisualization.Charting.AxisType.Primary
        End Select

        Select Case cmbYAxisValueType.SelectedItem.ToString
            Case "Auto"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Auto
            Case "Date"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Date
            Case "DateTime"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.DateTime
            Case "DateTimeOffset"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.DateTimeOffset
            Case "Double"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Double
            Case "Int32"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Int32
            Case "Int64"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Int64
            Case "Single"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Single
            Case "String"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.String
            Case "Time"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Time
            Case "UInt32"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.UInt32
            Case "UInt64"
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.UInt64
            Case Else
                Chart1.Series(SeriesName).YValueType = DataVisualization.Charting.ChartValueType.Auto
        End Select


        'Save custom attributes:
        If DataGridView2.Rows(0).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("LabelValueType", "Close")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("LabelValueType", DataGridView2.Rows(0).Cells(2).Value)
        End If
        If DataGridView2.Rows(1).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("MaxPixelPointWidth", "0")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("MaxPixelPointWidth", DataGridView2.Rows(1).Cells(2).Value)
        End If
        If DataGridView2.Rows(2).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("MinPixelPointWidth", "0")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("MinPixelPointWidth", DataGridView2.Rows(2).Cells(2).Value)
        End If
        If DataGridView2.Rows(3).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("OpenCloseStyle", "Line")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("OpenCloseStyle", DataGridView2.Rows(3).Cells(2).Value)
        End If
        If DataGridView2.Rows(4).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointDepth", "0")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointDepth", DataGridView2.Rows(4).Cells(2).Value)
        End If
        If DataGridView2.Rows(5).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", "0")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", DataGridView2.Rows(5).Cells(2).Value)
        End If
        If DataGridView2.Rows(6).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointWidth", "0")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("PixelPointWidth", DataGridView2.Rows(6).Cells(2).Value)
        End If
        If DataGridView2.Rows(7).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("PointWidth", "0.8")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("PointWidth", DataGridView2.Rows(7).Cells(2).Value)
        End If
        If DataGridView2.Rows(8).Cells(2).Value = "" Then
            Chart1.Series(SeriesName).SetCustomProperty("ShowOpenClose", "Both")
        Else
            Chart1.Series(SeriesName).SetCustomProperty("ShowOpenClose", DataGridView2.Rows(8).Cells(2).Value)
        End If



        'If DataGridView2.Rows(0).Cells(2).Value = "" Then
        '    Chart1.Series(SeriesName).SetCustomProperty("EmptyPointValue", "Average")
        'Else
        '    Chart1.Series(SeriesName).SetCustomProperty("EmptyPointValue", DataGridView2.Rows(0).Cells(2).Value)
        'End If
        'If DataGridView2.Rows(1).Cells(2).Value = "" Then
        '    Chart1.Series(SeriesName).SetCustomProperty("LabelStyle", "Auto")
        'Else
        '    Chart1.Series(SeriesName).SetCustomProperty("LabelStyle", DataGridView2.Rows(1).Cells(2).Value)
        'End If
        'If DataGridView2.Rows(2).Cells(2).Value = "" Then
        '    Chart1.Series(SeriesName).SetCustomProperty("PixelPointDepth", "1")
        'Else
        '    Chart1.Series(SeriesName).SetCustomProperty("PixelPointDepth", DataGridView2.Rows(2).Cells(2).Value)
        'End If
        'If DataGridView2.Rows(3).Cells(2).Value = "" Then
        '    Chart1.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", "1")
        'Else
        '    Chart1.Series(SeriesName).SetCustomProperty("PixelPointGapDepth", DataGridView2.Rows(3).Cells(2).Value)
        'End If

        'NOTE: THE FOLLOWING IS NOT USED FOR A POINT CHART:
        'If DataGridView2.Rows(4).Cells(2).Value = "" Then
        '    Chart1.Series(SeriesName).SetCustomProperty("ShowMarkerLines", "True")
        'Else
        '    Chart1.Series(SeriesName).SetCustomProperty("ShowMarkerLines", DataGridView2.Rows(4).Cells(2).Value)
        'End If

        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Area
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Bar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.BoxPlot
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Bubble
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Candlestick
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Column
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Doughnut
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.ErrorBar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.FastLine
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.FastPoint
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Funnel
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Kagi
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Line
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Pie
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Point
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.PointAndFigure
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Polar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Pyramid
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Radar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Range
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.RangeBar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.RangeColumn
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Renko
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Spline
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.SplineArea
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.SplineRange
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedArea
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedArea100
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedBar
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedBar100
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedColumn
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StackedColumn100
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.StepLine
        Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.Stock
        'Chart1.Series(SeriesName).ChartType = DataVisualization.Charting.SeriesChartType.ThreeLineBreak

        'Chart1.Series(SeriesName).ChartArea = "Area1"



        'Chart1.Series(SeriesName).MarkerStyle = DataVisualization.Charting.MarkerStyle.Circle
        'Chart1.Series(SeriesName).MarkerBorderColor = Color.Black
        'Chart1.Series(SeriesName).BorderWidth = 2
        'Chart1.Series(SeriesName).BorderColor = Color.Blue



        'Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Auto
        'Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Date
        'Chart1.Series(SeriesName).XValueType = DataVisualization.Charting.ChartValueType.Date

        'Check the Series colors:
        If Chart1.Series(SeriesName).MarkerBorderColor = Color.FromArgb(0) Then Chart1.Series(SeriesName).MarkerBorderColor = Color.Black
        If Chart1.Series(SeriesName).MarkerColor = Color.FromArgb(0) Then Chart1.Series(SeriesName).MarkerColor = Color.Gray
        If Chart1.Series(SeriesName).Color = Color.FromArgb(0) Then Chart1.Series(SeriesName).Color = Color.Blue

    End Sub

    Private Sub btnAddSeries_Click(sender As Object, e As EventArgs) Handles btnAddSeries.Click
        'Add a new series to the Chart:

        If txtSeriesRecordNo.Text = "0" Then
            'No Point Chart data has been loaded.
            'Check if there is a default Series already in the Chart control:
            If Chart1.Series.Count > 0 Then
                txtSeriesRecordNo.Text = "1"
                txtNSeriesRecords.Text = Chart1.Series.Count
                txtSeriesName.Text = Chart1.Series(0).Name
            Else
                'To Do!!!

            End If

        Else 'Add a new Series to the Chart control:
            'Dim NewSeriesNo As Integer = Chart1.Series.Count
            Dim NewSeriesNo As Integer = Chart1.Series.Count + 1

            Dim NewSeriesName As String = "Series" & NewSeriesNo
            Chart1.Series.Add(NewSeriesName)

            txtSeriesName.Text = NewSeriesName
            txtSeriesRecordNo.Text = NewSeriesNo + 1
            txtNSeriesRecords.Text = Chart1.Series.Count
        End If
    End Sub

    Private Sub UpdateSeriesTabSettings()
        'Update the Series tab settings from the Chart control and ChartInfo.

        Dim NSeries As Integer = Chart1.Series.Count
        'txtNSeriesRecords.Text = Chart1.Series.Count
        txtNSeriesRecords.Text = NSeries

        'If txtSeriesRecordNo.Text.Trim = "" Then
        '    txtSeriesRecordNo.Text = "1"
        'End If

        'Dim SeriesNo As Integer = Val(txtSeriesRecordNo.Text) - 1
        Dim SeriesNo As Integer

        'Dim SeriesName As String = Chart1.Series(SeriesNo).Name
        Dim SeriesName As String

        If NSeries = 0 Then
            txtSeriesRecordNo.Text = "0"
            SeriesNo = 0
            SeriesName = ""
            txtSeriesName.Text = ""
            cmbXValues.SelectedIndex = 0
        Else
            If txtSeriesRecordNo.Text.Trim = "" Then
                txtSeriesRecordNo.Text = "1"
            End If
            'SeriesNo = 1
            SeriesNo = Val(txtSeriesRecordNo.Text)
            If SeriesNo < 1 Then
                SeriesNo = 1
                'If SeriesNo > NSeries Then
                '    SeriesNo = NSeries
                'End If
            End If
            If SeriesNo > NSeries Then SeriesNo = NSeries
            txtSeriesRecordNo.Text = SeriesNo
            'txtSeriesName.Text = Chart1.Series(0).Name
            'SeriesName = Chart1.Series(SeriesNo).Name
            SeriesName = Chart1.Series(SeriesNo - 1).Name
            txtSeriesName.Text = SeriesName

            'Update list of areas in Series tab:
            cmbChartArea.Items.Clear()
            For Each item In Chart1.ChartAreas
                cmbChartArea.Items.Add(item.Name)
                'Message.Add("Adding Chare Area: " & item.Name & vbCrLf)
            Next

            If SeriesName = "" Then
                Message.AddWarning("The Series Name is blank" & vbCrLf)
                'ElseIf ChartInfo.dictFields.ContainsKey(SeriesName) Then
            ElseIf ChartInfo.dictSeriesInfo.ContainsKey(SeriesName) Then
                'Apply X Values Field Name:
                'cmbXValues.SelectedIndex = cmbXValues.FindStringExact(ChartInfo.dictFields(SeriesName).XValuesFieldName)
                cmbXValues.SelectedIndex = cmbXValues.FindStringExact(ChartInfo.dictSeriesInfo(SeriesName).XValuesFieldName)

                cmbXAxisType.SelectedIndex = cmbXAxisType.FindStringExact(Chart1.Series(SeriesName).XAxisType.ToString)
                cmbXAxisValueType.SelectedIndex = cmbXAxisValueType.FindStringExact(Chart1.Series(SeriesName).XValueType.ToString)

                'cmbChartArea.SelectedIndex = cmbChartArea.FindStringExact(ChartInfo.dictSeriesInfo(SeriesName).ChartArea)
                cmbChartArea.SelectedIndex = cmbChartArea.FindStringExact(Chart1.Series(SeriesName).ChartArea)
                'Apply Y Values selections:
                For I = 1 To cboFieldSelections.Items.Count
                    'If PointChart.YValuesFieldName = cboFieldSelections.Items(I - 1) Then
                    'If ChartInfo.dictFields(SeriesName).YValuesFieldName = cboFieldSelections.Items(I - 1) Then
                    'If ChartInfo.dictSeriesInfo(SeriesName).YValuesFieldName = cboFieldSelections.Items(I - 1) Then
                    '    DataGridView1.Rows(0).Cells(1).Value = cboFieldSelections.Items(I - 1)
                    'End If
                    If ChartInfo.dictSeriesInfo(SeriesName).YValuesHighFieldName = cboFieldSelections.Items(I - 1) Then
                        DataGridView1.Rows(0).Cells(1).Value = cboFieldSelections.Items(I - 1)
                    End If
                    If ChartInfo.dictSeriesInfo(SeriesName).YValuesLowFieldName = cboFieldSelections.Items(I - 1) Then
                        DataGridView1.Rows(1).Cells(1).Value = cboFieldSelections.Items(I - 1)
                    End If
                    If ChartInfo.dictSeriesInfo(SeriesName).YValuesOpenFieldName = cboFieldSelections.Items(I - 1) Then
                        DataGridView1.Rows(2).Cells(1).Value = cboFieldSelections.Items(I - 1)
                    End If
                    If ChartInfo.dictSeriesInfo(SeriesName).YValuesCloseFieldName = cboFieldSelections.Items(I - 1) Then
                        DataGridView1.Rows(3).Cells(1).Value = cboFieldSelections.Items(I - 1)
                    End If
                Next

                cmbYAxisType.SelectedIndex = cmbYAxisType.FindStringExact(Chart1.Series(SeriesName).YAxisType.ToString)
                cmbYAxisValueType.SelectedIndex = cmbYAxisValueType.FindStringExact(Chart1.Series(SeriesName).YValueType.ToString)

                'Apply Custom Attributes selections:
                If Chart1.Series(SeriesName).GetCustomProperty("LabelValueType") <> "" Then
                    DataGridView2.Rows(0).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("LabelValueType")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("MaxPixelPointWidth") <> "" Then
                    DataGridView2.Rows(1).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("MaxPixelPointWidth")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("MinPixelPointWidth") <> "" Then
                    DataGridView2.Rows(2).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("MinPixelPointWidth")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("OpenCloseStyle") <> "" Then
                    DataGridView2.Rows(3).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("OpenCloseStyle")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("PixelPointDepth") <> "" Then
                    DataGridView2.Rows(4).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointDepth")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("PixelPointGapDepth") <> "" Then
                    DataGridView2.Rows(5).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointGapDepth")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("PixelPointWidth") <> "" Then
                    DataGridView2.Rows(6).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointWidth")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("PointWidth") <> "" Then
                    DataGridView2.Rows(7).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PointWidth")
                End If
                If Chart1.Series(SeriesName).GetCustomProperty("ShowOpenClose") <> "" Then
                    DataGridView2.Rows(8).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("ShowOpenClose")
                End If



                'If Chart1.Series(SeriesName).GetCustomProperty("EmptyPointValue") <> "" Then
                '    DataGridView2.Rows(0).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("EmptyPointValue")
                'End If
                'If Chart1.Series(SeriesName).GetCustomProperty("LabelStyle") <> "" Then
                '    DataGridView2.Rows(1).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("LabelStyle")
                'End If
                'DataGridView2.Rows(2).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointDepth")
                'DataGridView2.Rows(3).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("PixelPointGapDepth")

                ''DataGridView2.Rows(4).Cells(2).Value = Chart1.Series(SeriesName).GetCustomProperty("ShowMarkerLines") 'NOT USED IN A POINT CHART
            Else
                Message.AddWarning("The Series Name is not in the Chart Info dictionary: " & SeriesName & vbCrLf)
            End If
        End If


        ''Apply Custom Attributes selections:
        ''LabelValueType (High, Low, Open, Close)
        ''If PointChart.EmptyPointValue <> "" Then
        'If Chart1.Series(SeriesName).CustomProperties("EmptyPointValue") <> "" Then
        '    'DataGridView2.Rows(0).Cells(2).Value = PointChart.EmptyPointValue
        '    DataGridView2.Rows(0).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("EmptyPointValue")
        'End If
        ''If PointChart.LabelStyle <> "" Then
        'If Chart1.Series(SeriesName).CustomProperties("LabelStyle") <> "" Then
        '    'DataGridView2.Rows(1).Cells(2).Value = PointChart.LabelStyle
        '    DataGridView2.Rows(1).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("LabelStyle")
        'End If
        ''PixelPointDepth (Any integer > 0)
        ''DataGridView2.Rows(2).Cells(2).Value = PointChart.PixelPointDepth
        'DataGridView2.Rows(2).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("PixelPointDepth")
        ''PixelPointGapDepth (Any integer > 0)
        ''DataGridView2.Rows(3).Cells(2).Value = PointChart.PixelPointGapDepth
        'DataGridView2.Rows(3).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("PixelPointGapDepth")

        'DataGridView2.Rows(4).Cells(2).Value = Chart1.Series(SeriesName).CustomProperties("ShowMarkerLines")

    End Sub

#End Region 'Chart Series Tab -------------------------------------------------------------------------------------------------------------------------------------------


#Region " Chart Areas Tab" '=============================================================================================================================================

    Private Sub UpdateAreasTabSettings()
        'Update the Areas tab settings from the Chart control.

        Dim NAreas As Integer = Chart1.ChartAreas.Count
        txtNAreaRecords.Text = NAreas

        Dim AreaNo As Integer
        Dim AreaName As String

        If NAreas = 0 Then
            txtAreaRecordNo.Text = "0"
            AreaNo = 0
            AreaName = ""
            txtAreaName.Text = ""

        Else
            If txtAreaRecordNo.Text.Trim = "" Then
                txtAreaRecordNo.Text = "1"
            End If

            'AreaNo = Val(txtAreaRecordNo.Text)
            AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.
            'If AreaNo < 1 Then
            '    AreaNo = 1
            'End If
            If AreaNo < 0 Then
                AreaNo = 0
            End If
            'If AreaNo > NAreas Then AreaNo = NAreas
            If AreaNo + 1 > NAreas Then AreaNo = NAreas - 1
            'txtAreaRecordNo.Text = AreaNo
            'txtAreaRecordNo.Text = AreaNo + 1
            ShowArea(AreaNo)

        End If

    End Sub

    Private Sub ApplyAreasTabSettings()
        'Apply the settings in the Areas tab to the Chart control.

        Dim AreaNo As Integer = Val(txtAreaRecordNo.Text)
        Dim AreaCount As Integer = Chart1.ChartAreas.Count

        If AreaNo - 1 > AreaCount Then
            Message.AddWarning("The area number is larger than to number of areas in the chart!" & vbCrLf)
            Exit Sub
        End If

        Dim AreaName As String = txtAreaName.Text.Trim
        Chart1.ChartAreas(AreaNo - 1).Name = AreaName

        'If ChartInfo.dictAreas.ContainsKey(AreaName) Then
        If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
            'AreaName is already in the dictionary of Area Auto settings.
        Else
            ChartInfo.dictAreaInfo.Add(AreaName, New AreaInfo)
        End If


        'Message.Add("1- Chart1.ChartAreas(0).AxisX.Title = " & Chart1.ChartAreas(0).AxisX.Title & vbCrLf) 'NaN



        'X Axis: -------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisX.Title = txtXAxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisX.TitleFont = txtXAxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisX.TitleForeColor = txtXAxisTitle.ForeColor
        If cmbXAxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisX.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbXAxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisX.LabelStyle.Format = txtXAxisLabelStyleFormat.Text


        If txtXAxisMin.Text.Trim = "" Then
            chkXAxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum = True
        ElseIf chkXAxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.Minimum = Val(txtXAxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum = False
        End If
        'Message.Add("2- Chart1.ChartAreas(0).AxisX.Minimum = " & Chart1.ChartAreas(0).AxisX.Minimum & vbCrLf) 'NaN

        If txtXAxisMax.Text.Trim = "" Then
            chkXAxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum = True
        ElseIf chkXAxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.Maximum = Val(txtXAxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum = False
        End If
        'Message.Add("3- Chart1.ChartAreas(0).AxisX.Maximum = " & Chart1.ChartAreas(0).AxisX.Maximum & vbCrLf)

        If txtXAxisAnnotInt.Text.Trim = "" Then
            chkXAxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkXAxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.Interval = Val(txtXAxisAnnotInt.Text)
        End If
        'Message.Add("4- Chart1.ChartAreas(0).AxisX.Interval = " & Chart1.ChartAreas(0).AxisX.Interval & vbCrLf)

        If txtXAxisMajGridInt.Text.Trim = "" Then
            chkXAxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval = True
        ElseIf chkXAxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX.MajorGrid.Interval = Val(txtXAxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval = False
        End If
        'Message.Add("5- Chart1.ChartAreas(0).AxisX.MajorGrid.Interval = " & Chart1.ChartAreas(0).AxisX.MajorGrid.Interval & vbCrLf)

        If chkXAxisScrollBar.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Enabled = True
            Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Size = 16
        End If


        'X2 Axis: ------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisX2.Title = txtX2AxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleFont = txtX2AxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleForeColor = txtX2AxisTitle.ForeColor
        If cmbX2AxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisX2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbX2AxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisX2.LabelStyle.Format = txtX2AxisLabelStyleFormat.Text

        If txtX2AxisMin.Text.Trim = "" Then
            chkX2AxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum = True
        ElseIf chkX2AxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Minimum = Val(txtX2AxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum = False
        End If

        If txtX2AxisMax.Text.Trim = "" Then
            chkX2AxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum = True
        ElseIf chkX2AxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Maximum = Val(txtX2AxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum = False
        End If

        If txtX2AxisAnnotInt.Text.Trim = "" Then
            chkX2AxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkX2AxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisX2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.Interval = Val(txtX2AxisAnnotInt.Text)
        End If

        If txtX2AxisMajGridInt.Text.Trim = "" Then
            chkX2AxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval = True
        ElseIf chkX2AxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisX2.MajorGrid.Interval = Val(txtX2AxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval = False
        End If

        'Y Axis: -------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisY.Title = txtYAxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisY.TitleFont = txtYAxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisY.TitleForeColor = txtYAxisTitle.ForeColor
        If cmbYAxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisY.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbYAxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisY.LabelStyle.Format = txtYAxisLabelStyleFormat.Text

        If txtYAxisMin.Text.Trim = "" Then
            chkYAxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum = True
        ElseIf chkYAxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.Minimum = Val(txtYAxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum = False
        End If

        If txtYAxisMax.Text.Trim = "" Then
            chkYAxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum = True
        ElseIf chkYAxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.Maximum = Val(txtYAxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum = False
        End If

        If txtYAxisAnnotInt.Text.Trim = "" Then
            chkYAxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkYAxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.Interval = Val(txtYAxisAnnotInt.Text)
        End If

        If txtYAxisMajGridInt.Text.Trim = "" Then
            chkYAxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval = True
        ElseIf chkYAxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY.MajorGrid.Interval = Val(txtYAxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval = False
        End If




        'Y2 Axis: ------------------------------------------------------------------------
        Chart1.ChartAreas(AreaNo - 1).AxisY2.Title = txtY2AxisTitle.Text
        Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleFont = txtY2AxisTitle.Font
        Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleForeColor = txtY2AxisTitle.ForeColor
        If cmbY2AxisTitleAlignment.SelectedItem IsNot Nothing Then Chart1.ChartAreas(AreaNo - 1).AxisY2.TitleAlignment = [Enum].Parse(GetType(StringAlignment), cmbY2AxisTitleAlignment.SelectedItem.ToString)
        Chart1.ChartAreas(AreaNo - 1).AxisY2.LabelStyle.Format = txtY2AxisLabelStyleFormat.Text


        If txtY2AxisMin.Text.Trim = "" Then
            chkY2AxisAutoMin.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum = True
        ElseIf chkY2AxisAutoMin.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum = Double.NaN 'Auto minimum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Minimum = Val(txtY2AxisMin.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum = False
        End If

        If txtY2AxisMax.Text.Trim = "" Then
            chkY2AxisAutoMax.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum = True
        ElseIf chkY2AxisAutoMax.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum = Double.NaN 'Auto maximum.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Maximum = Val(txtY2AxisMax.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum = False
        End If

        If txtY2AxisAnnotInt.Text.Trim = "" Then
            chkY2AxisAutoAnnotInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        ElseIf chkY2AxisAutoAnnotInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = 0 'Zero indicates Auto mode.
            Chart1.ChartAreas(AreaNo - 1).AxisY2.IntervalAutoMode = DataVisualization.Charting.IntervalAutoMode.VariableCount
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.Interval = Val(txtY2AxisAnnotInt.Text)
        End If

        If txtY2AxisMajGridInt.Text.Trim = "" Then
            chkY2AxisAutoMajGridInt.Checked = True
            Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval = True
        ElseIf chkY2AxisAutoMajGridInt.Checked Then
            Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval = Double.NaN 'Indicates Not Set - use Axis Interval value.
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval = True
        Else
            Chart1.ChartAreas(AreaNo - 1).AxisY2.MajorGrid.Interval = Val(txtY2AxisMajGridInt.Text)
            ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval = False
        End If


    End Sub

    Private Sub btnAddArea_Click(sender As Object, e As EventArgs) Handles btnAddArea.Click
        'Add a new area to the Chart:

        If txtAreaRecordNo.Text = "0" Then
            'No Point Chart data hase been loaded.
            'Check if there is a default Area already in the Chart control:
            If Chart1.ChartAreas.Count > 0 Then
                txtAreaRecordNo.Text = "1"
                txtNAreaRecords.Text = Chart1.ChartAreas.Count
                txtAreaName.Text = Chart1.ChartAreas(0).Name
            Else
                'To Do!!!
            End If
        Else 'Add a new Area to the Chart control:
            'Dim NewAreaNo As Integer = Chart1.ChartAreas.Count
            Dim NewAreaNo As Integer = Chart1.ChartAreas.Count + 1

            'Dim NewAreaName As String = "Area" & NewAreaNo
            Dim NewAreaName As String = "ChartArea" & NewAreaNo
            Chart1.ChartAreas.Add(NewAreaName)

            txtAreaName.Text = NewAreaName
            txtAreaRecordNo.Text = NewAreaNo + 1
            txtNAreaRecords.Text = Chart1.ChartAreas.Count
        End If

    End Sub

    Private Sub btnApplyAreaSettings_Click(sender As Object, e As EventArgs) Handles btnApplyAreaSettings.Click
        ApplyAreasTabSettings()
    End Sub

    Private Sub btnXAxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnXAxisTitleFont.Click
        FontDialog1.Font = txtXAxisTitle.Font
        FontDialog1.ShowDialog()
        txtXAxisTitle.Font = FontDialog1.Font
    End Sub
    Private Sub btnXAxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnXAxisTitleColor.Click
        ColorDialog1.Color = txtXAxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtXAxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub txtXAxisZoomInterval_LostFocus(sender As Object, e As EventArgs) Handles txtXAxisZoomInterval.LostFocus
        'Update ZoomTo value using the ZoomFrom and ZoomInterval values:

        txtXAxisZoomTo.Text = Val(txtXAxisZoomFrom.Text) + Val(txtXAxisZoomInterval.Text)
    End Sub

    Private Sub btnZoomOK_Click(sender As Object, e As EventArgs) Handles btnZoomOK.Click
        'Zoom to the X Axis range shown:

        Dim ZoomStart As Double
        Dim ZoomEnd As Double

        If txtXAxisZoomFrom.Text = "" Then

        Else
            ZoomStart = Val(txtXAxisZoomFrom.Text)
            If txtXAxisZoomTo.Text = "" Then

            Else
                ZoomEnd = Val(txtXAxisZoomTo.Text)
                Dim AreaNo As Integer = Val(txtAreaRecordNo.Text)
                Dim AreaCount As Integer = Chart1.ChartAreas.Count

                If AreaNo - 1 > AreaCount Then
                    Message.AddWarning("The area number is larger than to number of areas in the chart!" & vbCrLf)
                    Exit Sub
                End If

                Dim AreaName As String = txtAreaName.Text.Trim
                If chkXAxisScrollBar.Checked Then
                    Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Enabled = True
                    Chart1.ChartAreas(AreaNo - 1).AxisX.ScrollBar.Size = 16
                End If
                Chart1.ChartAreas(AreaNo - 1).AxisX.ScaleView.Zoom(ZoomStart, ZoomEnd)
            End If
        End If

    End Sub

    Private Sub btnX2AxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnX2AxisTitleColor.Click
        ColorDialog1.Color = txtX2AxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtX2AxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnYAxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnYAxisTitleColor.Click
        ColorDialog1.Color = txtYAxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtYAxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnY2AxisTitleColor_Click(sender As Object, e As EventArgs) Handles btnY2AxisTitleColor.Click
        ColorDialog1.Color = txtY2AxisTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtY2AxisTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub btnX2AxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnX2AxisTitleFont.Click
        FontDialog1.Font = txtX2AxisTitle.Font
        FontDialog1.ShowDialog()
        txtX2AxisTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnYAxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnYAxisTitleFont.Click
        FontDialog1.Font = txtYAxisTitle.Font
        FontDialog1.ShowDialog()
        txtYAxisTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnY2AxisTitleFont_Click(sender As Object, e As EventArgs) Handles btnY2AxisTitleFont.Click
        FontDialog1.Font = txtY2AxisTitle.Font
        FontDialog1.ShowDialog()
        txtY2AxisTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnDeleteArea_Click(sender As Object, e As EventArgs) Handles btnDeleteArea.Click
        'Delete the selected chart area:

        Dim AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.

        If Chart1.ChartAreas.Count > 1 Then
            Chart1.ChartAreas.RemoveAt(AreaNo)
            UpdateAreasTabSettings()
        Else
            'Only one chart area. This should be edited rather than deleted.
        End If
    End Sub

    Private Sub btnPrevArea_Click(sender As Object, e As EventArgs) Handles btnPrevArea.Click
        'Show the previous Area:

        Dim AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.

        If AreaNo = 0 Then
            'Already at the first Area.
        Else
            'Show the previous area:
            AreaNo = AreaNo - 1
            'txtAreaRecordNo.Text = AreaNo + 1
            ShowArea(AreaNo)
        End If

    End Sub

    Private Sub btnNextArea_Click(sender As Object, e As EventArgs) Handles btnNextArea.Click
        'Show the next area:

        Dim AreaNo = Val(txtAreaRecordNo.Text) - 1 'Zero-based area number.

        If AreaNo + 1 >= Chart1.ChartAreas.Count Then
            'Already at the last area.
        Else
            'Snow the next area:
            AreaNo = AreaNo + 1
            ShowArea(AreaNo)
        End If

    End Sub

    Private Sub ShowArea(ByVal AreaNo As Integer)
        'Show the Area for AreaNo (Zero-based index).

        If AreaNo < 0 Then
            'AreaNo is too small!
        ElseIf AreaNo + 1 > Chart1.ChartAreas.Count Then
            'AreaNo is too large!
        Else
            txtAreaRecordNo.Text = AreaNo + 1
            Dim AreaName As String = Chart1.ChartAreas(AreaNo).Name
            txtAreaName.Text = AreaName

            txtXAxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisX.Title
            txtX2AxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisX2.Title
            txtYAxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisY.Title
            txtY2AxisTitle.Text = Chart1.ChartAreas(AreaNo).AxisY2.Title

            txtXAxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisX.TitleFont
            txtX2AxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisX2.TitleFont
            txtYAxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisY.TitleFont
            txtY2AxisTitle.Font = Chart1.ChartAreas(AreaNo).AxisY2.TitleFont

            txtXAxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisX.TitleForeColor
            txtX2AxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisX2.TitleForeColor
            txtYAxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisY.TitleForeColor
            txtY2AxisTitle.ForeColor = Chart1.ChartAreas(AreaNo).AxisY2.TitleForeColor

            cmbXAxisTitleAlignment.SelectedIndex = cmbXAxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisX.TitleAlignment.ToString)
            cmbX2AxisTitleAlignment.SelectedIndex = cmbX2AxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisX2.TitleAlignment.ToString)
            cmbYAxisTitleAlignment.SelectedIndex = cmbYAxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisY.TitleAlignment.ToString)
            cmbY2AxisTitleAlignment.SelectedIndex = cmbY2AxisTitleAlignment.FindStringExact(Chart1.ChartAreas(AreaNo).AxisY2.TitleAlignment.ToString)

            'Axis Minimum values: --------------------------------------------
            If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
                chkXAxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMinimum
                txtXAxisMin.Text = Chart1.ChartAreas(AreaNo).AxisX.Minimum
                txtXAxisZoomFrom.Text = Chart1.ChartAreas(AreaNo).AxisX.Minimum
                chkX2AxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMinimum
                txtX2AxisMin.Text = Chart1.ChartAreas(AreaNo).AxisX2.Minimum
                chkYAxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMinimum
                txtYAxisMin.Text = Chart1.ChartAreas(AreaNo).AxisY.Minimum
                chkY2AxisAutoMin.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMinimum
                txtY2AxisMin.Text = Chart1.ChartAreas(AreaNo).AxisY2.Minimum
            End If


            'Axis Maximum values: -----------------------------------------
            If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
                chkXAxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMaximum
                txtXAxisMax.Text = Chart1.ChartAreas(AreaNo).AxisX.Maximum
                txtXAxisZoomTo.Text = Chart1.ChartAreas(AreaNo).AxisX.Maximum
                txtXAxisZoomInterval.Text = Chart1.ChartAreas(AreaNo).AxisX.Maximum - Chart1.ChartAreas(AreaNo).AxisX.Minimum
                chkX2AxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMaximum
                txtX2AxisMax.Text = Chart1.ChartAreas(AreaNo).AxisX2.Maximum
                chkYAxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMaximum
                txtYAxisMax.Text = Chart1.ChartAreas(AreaNo).AxisY.Maximum
                chkY2AxisAutoMax.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMaximum
                txtY2AxisMax.Text = Chart1.ChartAreas(AreaNo).AxisY2.Maximum
            End If

            'Axis Intervals: -------------------------------------------------------------
            If Chart1.ChartAreas(AreaNo).AxisX.Interval = 0 Then 'Auto mode.
                chkXAxisAutoAnnotInt.Checked = True
            Else
                chkXAxisAutoAnnotInt.Checked = False
            End If
            txtXAxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisX.Interval

            If Chart1.ChartAreas(AreaNo).AxisX2.Interval = 0 Then 'Auto mode.
                chkX2AxisAutoAnnotInt.Checked = True
            Else
                chkX2AxisAutoAnnotInt.Checked = False
            End If
            txtX2AxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisX2.Interval
            If Chart1.ChartAreas(AreaNo).AxisY.Interval = 0 Then 'Auto mode.
                chkYAxisAutoAnnotInt.Checked = True
            Else
                chkYAxisAutoAnnotInt.Checked = False
            End If
            txtYAxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisY.Interval
            If Chart1.ChartAreas(AreaNo).AxisY2.Interval = 0 Then 'Auto mode.
                chkY2AxisAutoAnnotInt.Checked = True
            Else
                chkY2AxisAutoAnnotInt.Checked = False
            End If
            txtY2AxisAnnotInt.Text = Chart1.ChartAreas(AreaNo).AxisY2.Interval

            'Axis Major Grid Intervals: -----------------------------------------------------------
            If ChartInfo.dictAreaInfo.ContainsKey(AreaName) Then
                chkXAxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoXAxisMajorGridInterval
                txtXAxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisX.MajorGrid.Interval
                chkX2AxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoX2AxisMajorGridInterval
                txtX2AxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisX2.MajorGrid.Interval
                chkYAxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoYAxisMajorGridInterval
                txtYAxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisY.MajorGrid.Interval
                chkY2AxisAutoMajGridInt.Checked = ChartInfo.dictAreaInfo(AreaName).AutoY2AxisMajorGridInterval
                txtY2AxisMajGridInt.Text = Chart1.ChartAreas(AreaNo).AxisY2.MajorGrid.Interval
            End If

            'Axis Label Style Format:
            txtXAxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisX.LabelStyle.Format
            txtX2AxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisX2.LabelStyle.Format
            txtYAxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisY.LabelStyle.Format
            txtY2AxisLabelStyleFormat.Text = Chart1.ChartAreas(AreaNo).AxisY2.LabelStyle.Format

            'Update list of areas in Series tab:
            cmbChartArea.Items.Clear()
            For Each item In Chart1.ChartAreas
                cmbChartArea.Items.Add(item.Name)
                'Message.Add("Adding Chare Area: " & item.Name & vbCrLf)
            Next
            Dim SeriesName As String = txtSeriesName.Text
            cmbChartArea.SelectedItem = cmbChartArea.FindStringExact(Chart1.Series(SeriesName).ChartArea)
        End If
    End Sub

    Private Sub btnChartTitleFont_Click_1(sender As Object, e As EventArgs) Handles btnChartTitleFont.Click
        'Edit chart title font
        FontDialog1.Font = txtChartTitle.Font
        FontDialog1.ShowDialog()
        txtChartTitle.Font = FontDialog1.Font
    End Sub

    Private Sub btnChartTitleColor_Click(sender As Object, e As EventArgs) Handles btnChartTitleColor.Click
        ColorDialog1.Color = txtChartTitle.ForeColor
        ColorDialog1.ShowDialog()
        txtChartTitle.ForeColor = ColorDialog1.Color
    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click

    End Sub



#End Region 'Chart Areas Tab --------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        'Clear the current chart:
        Chart1.ChartAreas.Clear()
        ChartInfo.dictAreaInfo.Clear()
        Chart1.ChartAreas.Add("ChartArea1")
        ChartInfo.dictAreaInfo.Add("ChartArea1", New AreaInfo)

        Chart1.Series.Clear()
        ChartInfo.dictSeriesInfo.Clear()
        Chart1.Series.Add("Series1")
        ChartInfo.dictSeriesInfo.Add("Series1", New SeriesInfo)

        UpdateChartSettingsTabs()
    End Sub

    Private Sub UpdateChartSettingsTabs()
        UpdateInputDataTabSettings()
        UpdateTitlesTabSettings()
        UpdateAreasTabSettings() 'Update Areas Tab before Series Tab. This will update the list of Chart Areas on the Series Tab.
        UpdateSeriesTabSettings()
    End Sub

    Private Sub Project_NewProjectCreated(ProjectPath As String) Handles Project.NewProjectCreated
        SendProjectInfo(ProjectPath) 'Send the path of the new project to the Network application. The new project will be added to the list of projects.
    End Sub

    Private Sub btnDelTitle_Click(sender As Object, e As EventArgs) Handles btnDelTitle.Click

    End Sub





#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class

Public Class clsSendMessageParams
    'Parameters used when sending a message using the Message Service.
    Public ProjectNetworkName As String
    Public ConnectionName As String
    Public Message As String
End Class
