Imports Microsoft.VisualBasic
Imports System.Configuration
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.IO


' Custom Exception Handling Code
' (c) 2008-2009 by Darin Higgins All Rights Reserved
' Based on a very stripped down version of an original concept
' which was itself based on a concept for
' VB6 Try-Catch style error handling published in VBPJ many moons ago.

' Portions adapted from articles by:
' Jeff Atwood (Unhandled Exception handling)
'    http://www.codeproject.com/KB/exception/ExceptionHandling.aspx
' Susan Abraham (Custom Exception Handling)
'    http://www.vbdotnetheaven.com/UploadFile/susanabraham/CustomExceptionHandlingvb11122005021337AM/CustomExceptionHandlingvb.aspx
' Scott Guthrie (sending email via System.Net.Mail)
'    http://weblogs.asp.net/scottgu/archive/2005/12/10/432854.aspx
' Michael  Nemtsev (Setting up a global error handler)
'    http://www.devnewsgroups.net/group/microsoft.public.dotnet.framework.windowsforms/topic63104.aspx
' The Scarms (Various Exception handling articles)
'    http://www.thescarms.com/dotnet/dotnetdebug.aspx
' 
' Please Note: Care has been taken to ensure that this code file can be essentially
' "dropped into" any project with only one dependency (on the LineMap.vb file). However, this does mean
' that there are some routines that might be redundant here.
' 
' Also, the line number resolution code depends on a Line number resource or LINEMAP file
' compiled with the GenerateLineMap utility


#Region " ExceptionDialog"
''' <summary>
''' The Exception Handler Dialog
'''
''' Generic user error dialog
'''
''' UI adapted from
'''
''' Alan Cooper's "About Face: The Essentials of User Interface Design"
''' Chapter VII, "The End of Errors", pages 423-440
'''
''' Original coding by Jeff Atwood
'''     http://www.codinghorror.com
''' 
''' Extension rework by Darin Higgins
'''    http://www.vbfengshui.com
'''
''' </summary>
''' <remarks></remarks>
''' <editHistory></editHistory>
Public Class ExceptionDialog

    Const PADSIZE As Integer = 10

    ''' <summary>
    ''' security-safe process.start wrapper
    ''' </summary>
    ''' <param name="strUrl"></param>
    ''' <remarks></remarks>
    Private Sub LaunchLink(ByVal strUrl As String)
        Try
            System.Diagnostics.Process.Start(strUrl)
        Catch ex As System.Security.SecurityException
            '-- do nothing; we can't launch without full trust.
            MsgBox("Unable to launch the link without full trust.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
        Catch ex As Exception
            MsgBox("Unable to launch the link at this time.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly)
        End Try
    End Sub


    Private Sub SizeBox(ByVal ctl As System.Windows.Forms.RichTextBox)
        Dim g As Graphics = Nothing
        Try
            '-- note that the height is taken as MAXIMUM, so size the label for maximum desired height!
            g = Graphics.FromHwnd(ctl.Handle)
            Dim objSizeF As SizeF = g.MeasureString(ctl.Text, ctl.Font, New SizeF(ctl.Width, ctl.Height))
            g.Dispose()
            ctl.Height = Convert.ToInt32(objSizeF.Height) + 5
        Catch ex As System.Security.SecurityException
            '-- do nothing; we can't set control sizes without full trust
        Finally
            If g IsNot Nothing Then g.Dispose()
        End Try
    End Sub


    Private Function DetermineDialogResult(ByVal strButtonText As String) As DialogResult
        '-- strip any accelerator keys we might have
        strButtonText = strButtonText.Replace("&", "")
        Select Case strButtonText.ToLower
            Case "abort"
                Return DialogResult.Abort
            Case "cancel"
                Return DialogResult.Cancel
            Case "ignore"
                Return DialogResult.Ignore
            Case "no"
                Return DialogResult.No
            Case "none"
                Return DialogResult.None
            Case "ok"
                Return DialogResult.OK
            Case "retry"
                Return DialogResult.Retry
            Case "yes"
                Return DialogResult.Yes
            Case "debug"
                '---- can't pass an arbitrary value back out
                '     so we have to do this funkiness
                Me.DebugClicked = True
                Return DialogResult.OK
        End Select
    End Function


    Private _DebugClicked As Boolean = False
    Public Property DebugClicked() As Boolean
        Get
            Return _DebugClicked
        End Get
        Set(ByVal value As Boolean)
            _DebugClicked = value
        End Set
    End Property


    Private Sub btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1.Click, btn2.Click, btn3.Click, btn4.Click
        Me.Close()
        Me.DialogResult = DetermineDialogResult(DirectCast(sender, Button).Text)
    End Sub


    Private Sub frmExceptionHandler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '-- make sure our window is on top
        Me.TopMost = True
        Me.TopMost = False

        '-- More >> has to be expanded
        Me.txtMore.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtMore.Visible = False

        '-- size the labels' height to accommodate the amount of text in them
        SizeBox(ScopeBox)
        SizeBox(ActionBox)
        SizeBox(ErrorBox)

        '-- now shift everything up
        lblScopeHeading.Top = ErrorBox.Top + ErrorBox.Height + PADSIZE
        ScopeBox.Top = lblScopeHeading.Top + lblScopeHeading.Height + PADSIZE

        lblActionHeading.Top = ScopeBox.Top + ScopeBox.Height + PADSIZE
        ActionBox.Top = lblActionHeading.Top + lblActionHeading.Height + PADSIZE

        lblMoreHeading.Top = ActionBox.Top + ActionBox.Height + PADSIZE
        btnMore.Top = lblMoreHeading.Top - 3
        btnMore.Tag = ">>"
        btnMore.Image = picUp2.Image

        Me.Height = btnMore.Top + btnMore.Height + PADSIZE + 65

        Me.DebugClicked = False
        Me.CenterToScreen()
    End Sub


    Private Sub btnMore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMore.Click
        If btnMore.Tag.ToString = ">>" Then
            Me.Height = Me.Height + 300
            With txtMore
                .Location = New System.Drawing.Point(lblMoreHeading.Left, lblMoreHeading.Top + lblMoreHeading.Height + PADSIZE)
                .Height = Me.ClientSize.Height - txtMore.Top - 45
                .Width = Me.ClientSize.Width - 2 * PADSIZE - txtMore.Left
                .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom _
                            Or AnchorStyles.Left Or AnchorStyles.Right
                .Visible = True
            End With
            btn3.Focus()
            btnMore.Tag = "<<"
            btnMore.Image = picDrop2.Image
        Else
            Me.SuspendLayout()
            btnMore.Tag = ">>"
            btnMore.Image = picUp2.Image
            Me.Height = btnMore.Top + btnMore.Height + PADSIZE + 65
            txtMore.Visible = False
            txtMore.Anchor = AnchorStyles.None
            Me.ResumeLayout()
        End If
    End Sub


    Private Sub ErrorBox_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs) Handles ErrorBox.LinkClicked
        LaunchLink(e.LinkText)
    End Sub


    Private Sub ScopeBox_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs) Handles ScopeBox.LinkClicked
        LaunchLink(e.LinkText)
    End Sub


    Private Sub ActionBox_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs) Handles ActionBox.LinkClicked
        LaunchLink(e.LinkText)
    End Sub


    ''' <summary>
    ''' If the user clicks here, copy the exception report to the clipboard
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CopyToClipboardLink_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyToClipboardLink.Click
        Dim sb = New StringBuilder
        sb.Append("Error:") : sb.AppendLine()
        sb.Append(ErrorBox.Text) : sb.AppendLine() : sb.AppendLine()
        sb.Append("Scope:") : sb.AppendLine()
        sb.Append(ScopeBox.Text) : sb.AppendLine() : sb.AppendLine()
        sb.Append("Action:") : sb.AppendLine()
        sb.Append(ActionBox.Text) : sb.AppendLine() : sb.AppendLine()
        sb.Append("Details:") : sb.AppendLine()
        sb.Append(txtMore.Text) : sb.AppendLine() : sb.AppendLine() : sb.AppendLine()

        Clipboard.SetText(sb.ToString)
    End Sub
End Class
#End Region


#Region " Exception Extension Methods"
Public Module ExceptionExtensions
    Private Const ERRCLASSNAME As String = "ExceptionExtensions"

    ''' <summary>
    ''' Enumeration of the Default buttons that
    ''' are available on the dialog
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DefaultErrorButtons
        [Default] = 0
        Button1 = 1
        Button2 = 2
        Button3 = 3
    End Enum


    Private _AppProduct As String = String.Empty
    Private _AppCompany As String = String.Empty
    Private _AppTitle As String = String.Empty
    Private _AppDescription As String = String.Empty
    Private _AppCodeBase As String = String.Empty
    Private _AppBuildDate As String = String.Empty
    Private _AppVersion As String = String.Empty
    Private _AppFullName As String = String.Empty
    Private _AppFileName As String = String.Empty


    ''' <summary>
    ''' Provides enhanced stack tracing output that includes line numbers, IL Position, and even column
    ''' numbers when PDB files are available
    ''' </summary>
    ''' <param name="thisException"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()> _
    Public Function EnhancedStackTrace(ByVal thisException As Exception) As String
        Return InternalEnhancedStackTrace(thisException)
    End Function


    ''' <summary>
    ''' Show a decent error dialog, all parms are optional and will be defaulted
    ''' </summary>
    ''' <param name="WhatHappened"></param>
    ''' <param name="HowUserAffected"></param>
    ''' <param name="WhatUserCanDo"></param>
    ''' <param name="MoreDetails"></param>
    ''' <param name="Exception"></param>
    ''' <param name="Buttons"></param>
    ''' <param name="Icon"></param>
    ''' <param name="DefaultButton"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()> _
    Public Function Show(ByVal Exception As Exception, _
            Optional ByVal WhatHappened As String = "", _
            Optional ByVal HowUserAffected As String = "", _
            Optional ByVal WhatUserCanDo As String = "", _
            Optional ByVal MoreDetails As String = "", _
            Optional ByVal Buttons As MessageBoxButtons = MessageBoxButtons.OK, _
            Optional ByVal Icon As MessageBoxIcon = MessageBoxIcon.Warning, _
            Optional ByVal DefaultButton As DefaultErrorButtons = DefaultErrorButtons.Default) As DialogResult

        Return ShowInternal(WhatHappened, HowUserAffected, WhatUserCanDo, MoreDetails, Exception, Buttons, Icon, DefaultButton)
    End Function


    ''' <summary>
    ''' internal method to show error dialog
    ''' </summary>
    ''' <param name="WhatHappened"></param>
    ''' <param name="HowUserAffected"></param>
    ''' <param name="WhatUserCanDo"></param>
    ''' <param name="MoreDetails"></param>
    ''' <param name="Buttons"></param>
    ''' <param name="Icon"></param>
    ''' <param name="DefaultButton"></param>
    ''' <returns></returns>
    Private Function ShowInternal( _
                Optional ByVal WhatHappened As String = "", _
                Optional ByVal HowUserAffected As String = "", _
                Optional ByVal WhatUserCanDo As String = "", _
                Optional ByVal MoreDetails As String = "", _
                Optional ByVal Exception As System.Exception = Nothing, _
                Optional ByVal Buttons As MessageBoxButtons = MessageBoxButtons.OK, _
                Optional ByVal Icon As MessageBoxIcon = MessageBoxIcon.Error, _
                Optional ByVal DefaultButton As DefaultErrorButtons = DefaultErrorButtons.Button1) As DialogResult

        '---- make sure we're retrieved all asm attributes
        '     if we already have, this is short-circuited
        GetAssemblyAttribs()

        '-- set default values, etc
        If Len(WhatHappened) = 0 Then WhatHappened = "An unexpected problem has a occurred while trying to process your request. The problem was reported as '(exception)'."
        If Len(HowUserAffected) = 0 Then HowUserAffected = "The last operation you attempted has failed and wasn't performed."
        If Len(WhatUserCanDo) = 0 Then WhatUserCanDo = "You should attempt the operation again. If that fails, unload and restart (product)."

        ProcessStrings(Exception, WhatHappened, HowUserAffected, WhatUserCanDo, MoreDetails)

        If WillLogToFile Then OnWriteLog(Exception, WhatHappened)

        With New ExceptionDialog
            .Text = ReplaceStringVals(Exception, .Text)
            .ErrorBox.Text = WhatHappened
            .ScopeBox.Text = HowUserAffected
            .ActionBox.Text = WhatUserCanDo
            .txtMore.Text = MoreDetails

            '-- determine what button text, visibility, and defaults are
            Dim btn1 As Button
            Dim btn2 As Button
            Dim btn3 As Button
            If IsIDE Then
                '---- button 4 is the debugger button in this case
                .btn4.Text = "&Debug"
                btn1 = .btn1
                btn2 = .btn2
                btn3 = .btn3
            Else
                '---- redirect the other buttons
                .btn1.Visible = False
                btn1 = .btn2
                btn2 = .btn3
                btn3 = .btn4
            End If
            Select Case Buttons
                Case MessageBoxButtons.AbortRetryIgnore
                    btn1.Text = "&Abort"
                    btn2.Text = "&Retry"
                    btn3.Text = "&Ignore"
                    .AcceptButton = btn2
                    .CancelButton = btn3
                Case MessageBoxButtons.OK
                    btn3.Text = "OK"
                    btn2.Visible = False
                    btn1.Visible = False
                    .AcceptButton = btn3
                Case MessageBoxButtons.OKCancel
                    btn3.Text = "Cancel"
                    btn2.Text = "OK"
                    btn1.Visible = False
                    .AcceptButton = btn2
                    .CancelButton = btn3
                Case MessageBoxButtons.RetryCancel
                    btn3.Text = "Cancel"
                    btn2.Text = "&Retry"
                    btn1.Visible = False
                    .AcceptButton = btn2
                    .CancelButton = btn3
                Case MessageBoxButtons.YesNo
                    btn3.Text = "&No"
                    btn2.Text = "&Yes"
                    btn1.Visible = False
                Case MessageBoxButtons.YesNoCancel
                    btn3.Text = "Cancel"
                    btn2.Text = "&No"
                    btn1.Text = "&Yes"
                    .CancelButton = btn3
            End Select

            '-- set the proper dialog icon
            Select Case Icon
                Case MessageBoxIcon.Error
                    .PictureBox1.Image = System.Drawing.SystemIcons.Error.ToBitmap
                Case MessageBoxIcon.Stop
                    .PictureBox1.Image = System.Drawing.SystemIcons.Error.ToBitmap
                Case MessageBoxIcon.Exclamation
                    .PictureBox1.Image = System.Drawing.SystemIcons.Exclamation.ToBitmap
                Case MessageBoxIcon.Information
                    .PictureBox1.Image = System.Drawing.SystemIcons.Information.ToBitmap
                Case MessageBoxIcon.Question
                    .PictureBox1.Image = System.Drawing.SystemIcons.Question.ToBitmap
                Case Else
                    .PictureBox1.Image = System.Drawing.SystemIcons.Error.ToBitmap
            End Select

            '-- override the default button
            '   Yeah, this is pretty nasty, what we're doing is making sure
            '   that if the called indicates btn2 should have focus, then
            '   the second button on screen will have focus, as opposed to
            '   always setting focus to btn2, which may not even be visible!
            Select Case DefaultButton
                Case DefaultErrorButtons.Button1
                    If btn1.Visible Then
                        .AcceptButton = btn1
                        btn1.TabIndex = 0
                    ElseIf btn2.Visible Then
                        .AcceptButton = btn2
                        btn2.TabIndex = 0
                    Else
                        .AcceptButton = btn3
                        btn3.TabIndex = 0
                    End If
                Case DefaultErrorButtons.Button2
                    If btn1.Visible Then
                        .AcceptButton = btn2
                        btn2.TabIndex = 0
                    ElseIf btn2.Visible Then
                        .AcceptButton = btn3
                        btn3.TabIndex = 0
                    End If
                Case DefaultErrorButtons.Button3
                    .AcceptButton = btn3
                    btn3.TabIndex = 0
            End Select

            '-- show the user our error dialog
            Dim r As DialogResult
            r = .ShowDialog()
            If .DebugClicked Then
                r = DirectCast(99, DialogResult)
                If IsIDE() Then
                    '---- if you've ended up here, 
                    '     just "step next" till you get back to your code
                    Stop
                End If
            End If

            '---- and return the result
            Return r
        End With
    End Function


    Private _WillIgnoreDebugErrors As Boolean = False
    ''' <summary>
    ''' Should Unhandled Exceptions be ignored during debugging
    ''' Normally, this will be false
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property WillIgnoreDebugErrors() As Boolean
        Get
            Return _WillIgnoreDebugErrors
        End Get
        Set(ByVal Value As Boolean)
            If Value <> _WillIgnoreDebugErrors Then
                _WillIgnoreDebugErrors = Value
                '---- reinit the handlers
                InstallHandlers()
            End If
        End Set
    End Property


    Private _WillDisplayDialog As Boolean = True
    ''' <summary>
    ''' Should Unhandled Errors display a dialog
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property WillDisplayUnhandledExceptionDialog() As Boolean
        Get
            Return _WillDisplayDialog
        End Get
        Set(ByVal Value As Boolean)
            _WillDisplayDialog = Value
        End Set
    End Property


    Private _WillDisplayHandledErrDialogs As Boolean = True
    ''' <summary>
    ''' Should Handled Error dialogs ever be displayed?
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property WillDisplayHandledExceptionDialogs() As Boolean
        Get
            Return _WillDisplayHandledErrDialogs
        End Get
        Set(ByVal Value As Boolean)
            _WillDisplayHandledErrDialogs = Value
        End Set
    End Property


    Private _WillKillAppOnUnhandledException As Boolean = False
    Public Property WillKillAppOnUnhandledException() As Boolean
        Get
            Return _WillKillAppOnUnhandledException
        End Get
        Set(ByVal Value As Boolean)
            _WillKillAppOnUnhandledException = Value
        End Set
    End Property


    Private _WillLogToFile As Boolean = True
    Public Property WillLogToFile() As Boolean
        Get
            Return _WillLogToFile
        End Get
        Set(ByVal Value As Boolean)
            _WillLogToFile = Value
        End Set
    End Property


    Private _WillLogToEventLog As Boolean = False
    Public Property WillLogToEventLog() As Boolean
        Get
            Return _WillLogToEventLog
        End Get
        Set(ByVal Value As Boolean)
            _WillLogToEventLog = Value
        End Set
    End Property


    ''' <summary>
    ''' Public property to simplify detection of running in the IDE
    ''' NOTE: if /DEBUG is present on the command line, this will force
    ''' the app into debugger mode, even if it's not in an IDE or debugger
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Private ReadOnly Property IsIDE() As Boolean
        Get
            If InStr(1, Command, "/debug", CompareMethod.Text) = 0 Then
                Return System.Diagnostics.Debugger.IsAttached
            Else
                Return True
            End If
        End Get
    End Property


    ''' <summary>
    ''' This is in a private routine for .NET security reasons
    ''' if this line of code is in a sub, the entire sub is tagged as full trust
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub KillProcess()
        System.Diagnostics.Process.GetCurrentProcess.Kill()
    End Sub


    ''' <summary>
    ''' Initialize Unhandled exception handlers
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Initialize()
        InstallHandlers()
    End Sub


    ''' <summary>
    ''' Install all unhandled Exception Handlers
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InstallHandlers()
        '-- we don't need an unhandled exception handler if we are running inside
        '-- the vs.net IDE; it is our "unhandled exception handler" in that case
        If WillIgnoreDebugErrors Then
            If IsIDE Then Return
        End If

        '-- track the parent assembly that set up error handling
        '-- need to call this NOW so we set it appropriately; otherwise
        '-- we may get the wrong assembly at exception time!
        ParentAssembly()

        '-- for winforms applications
        RemoveHandler Application.ThreadException, AddressOf UnhandledThreadExceptionHandler
        AddHandler Application.ThreadException, AddressOf UnhandledThreadExceptionHandler

        '-- for console applications
        RemoveHandler System.AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledAppExceptionHandler
        AddHandler System.AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledAppExceptionHandler
    End Sub


    ''' <summary>
    ''' handles Application.ThreadException event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub UnhandledThreadExceptionHandler(ByVal sender As System.Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
        GenericUnhandledExceptionHandler(e.Exception)
    End Sub


    ''' <summary>
    ''' handles AppDomain.CurrentDomain.UnhandledException event
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="args"></param>
    ''' <remarks></remarks>
    Private Sub UnhandledAppExceptionHandler(ByVal sender As System.Object, ByVal args As UnhandledExceptionEventArgs)
        Dim Exception As Exception = DirectCast(args.ExceptionObject, Exception)
        GenericUnhandledExceptionHandler(Exception)
    End Sub


    ''' <summary>
    ''' generic exception handler; the various specific handlers all call into this sub
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GenericUnhandledExceptionHandler(ByVal Ex As Exception)
        '-- turn the exception into an informative string

        If Not IsConsoleApp Then
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        End If

        '-- log this error to various locations
        Try
            If WillLogToFile Then OnWriteLog(Ex, "")

        Catch
            '---- Always catch anything else because any exceptions inside the UEH
            '     will cause the code to terminate immediately

        End Try

        If Not IsConsoleApp Then
            Cursor.Current = System.Windows.Forms.Cursors.Default
        End If

        '-- display message to the user
        If WillDisplayUnhandledExceptionDialog Then UnhandledExceptionToUI(Ex)

        If WillKillAppOnUnhandledException Then
            KillProcess()
            Application.Exit()
        End If
    End Sub


    ''' <summary>
    ''' Returns whether this app is a console app or not
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Private ReadOnly Property IsConsoleApp() As Boolean
        Get
            Return My.Application.GetType.BaseType.Name = "ConsoleApplicationBase"
        End Get
    End Property


    ''' <summary>
    ''' display a dialog to the user; otherwise we just terminate with no alert at all!
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UnhandledExceptionToUI(ByVal Ex As Exception)
        Dim HowUserAffected As String

        If WillKillAppOnUnhandledException Then
            HowUserAffected = "When you click OK, (app) will close."
        Else
            HowUserAffected = "The action you requested was not performed."
        End If

        If Not IsConsoleApp() Then
            '-- pop the dialog
            Ex.Show("There was an unexpected error in (app). This may be due to a programming issue or other unexpected issue.", _
                HowUserAffected, _
                "Restart (app), and try repeating your last action. Also try alternative methods of performing the same action.", _
                , _
                , _
                MessageBoxIcon.Stop)
        Else
            '-- note that writing to console pauses for ENTER
            '-- otherwise console window just terminates immediately
            ExceptionToConsole(Ex)
        End If
    End Sub


    ''' <summary>
    ''' turns exception into something an average user can hopefully
    ''' understand; still very technical
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <param name="bConsoleApp"></param>
    ''' <returns></returns>
    Private Function FormatExceptionForUser(ByVal Ex As Exception, ByVal bConsoleApp As Boolean) As String
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strBullet As String
        If bConsoleApp Then
            strBullet = "-"
        Else
            strBullet = "•"
        End If

        With objStringBuilder
            If Not bConsoleApp Then
                .Append("If you need immediate assistance, contact (contact).")
            End If
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)
            .Append("The following information about the error was automatically captured: ")
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)
            If WillLogToFile Then
                .Append(" ")
                .Append(strBullet)
                .Append(" ")
                'If rbLogToFileOK Then
                '    .Append("details were written to a text log at:")
                'Else
                '    .Append("details could NOT be written to the text log at:")
                'End If
                '.Append(Environment.NewLine)
                '.Append("   ")
                '.Append(rLogFullPath)
                .Append(Environment.NewLine)
            End If
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)
            .Append("Detailed error information follows:")
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)
            .Append(ExceptionToString(Ex))
        End With
        Return objStringBuilder.ToString
    End Function


    ''' <summary>
    ''' write an exception to the console
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ExceptionToConsole(ByVal Ex As Exception)
        Console.WriteLine("This application encountered an unexpected problem.")
        Console.WriteLine(FormatExceptionForUser(Ex, True))
        Console.WriteLine("The application must now terminate. Press ENTER to continue...")
        Console.ReadLine()
    End Sub


    Private _TakeScreenShotOK As Boolean
    ''' <summary>
    ''' take a desktop screenshot of our exception
    ''' note that this fires BEFORE the user clicks on the OK dismissing the crash dialog
    ''' so the crash dialog itself will not be displayed
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ExceptionToScreenshot()
        '-- note that screenshotname does NOT include the file type extension
        Try
            'rCachedErr.ScreenShotFiles = pTakeScreenShot()
            _TakeScreenShotOK = True

        Catch ex As Exception
            '---- failed to take screenshot
            _TakeScreenShotOK = False
        End Try
    End Sub


    ''' <summary>
    ''' Returns an array of strings containing the filename(s) of 
    ''' images taken for the screenshot
    ''' </summary>
    ''' <param name="windowToShoot"></param>
    ''' <returns></returns>
    Private Function TakeScreenShot(Optional ByVal windowToShoot As Control = Nothing) As List(Of String)
        Dim screensToProcess() As Screen
        Dim Files As List(Of String) = New List(Of String)

        If windowToShoot Is Nothing Then
            screensToProcess = Screen.AllScreens
        Else
            screensToProcess = New Screen() {Screen.FromControl(windowToShoot)}
        End If

        For i As Integer = 0 To screensToProcess.Length
            Dim thisFileName As String = String.Format("{0}\\{1}_ScreenShot_Screen{2}.png", System.IO.Path.GetTempPath(), _AppProduct, (i + 1))
            ScreenShot(screensToProcess(i), thisFileName)
            Files.Add(thisFileName)
        Next

        Return Files
    End Function


    ''' <summary>
    ''' Actually take a screenshot in png format for a specific screen
    ''' Save to the given filename
    ''' </summary>
    ''' <param name="screen"></param>
    ''' <param name="fileName"></param>
    ''' <remarks></remarks>
    Private Sub ScreenShot(ByVal screen As Screen, ByVal fileName As String)
        Dim imageSize As Size = New Size(screen.WorkingArea.Width, screen.WorkingArea.Height)

        Dim bitmap As Bitmap = New Bitmap(imageSize.Width, imageSize.Height, PixelFormat.Format24bppRgb)

        Dim g As Graphics = Graphics.FromImage(bitmap)
        g.CopyFromScreen(screen.WorkingArea.X, screen.WorkingArea.Y, 0, 0, imageSize)

        bitmap.Save(fileName, ImageFormat.Png)
    End Sub


    Private _LogToEventLogOK As Boolean
    ''' <summary>
    ''' write an exception to the Windows NT event log
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ExceptionToEventLog(ByVal Ex As Exception)
        Try
            System.Diagnostics.EventLog.WriteEntry( _
                System.AppDomain.CurrentDomain.FriendlyName, _
                Environment.NewLine & ExceptionToString(Ex), _
                EventLogEntryType.Error)
            _LogToEventLogOK = True
        Catch
            _LogToEventLogOK = False
        End Try
    End Sub


    ''' <summary>
    ''' translate exception object to string, with additional system info
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <returns></returns>
    Private Function ExceptionToString(ByVal Ex As Exception) As String
        Try
            Dim objStringBuilder As New System.Text.StringBuilder

            If Not (Ex.InnerException Is Nothing) Then
                '-- sometimes the original exception is wrapped in a more relevant outer exception
                '-- the detail exception is the "inner" exception
                '-- see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/exceptdotnet.asp
                With objStringBuilder
                    .Append("(Inner Exception)")
                    .Append(Environment.NewLine)
                    .Append(ExceptionToString(Ex.InnerException))
                    .Append(Environment.NewLine)
                    .Append("(Outer Exception)")
                    .Append(Environment.NewLine)
                End With
            End If
            With objStringBuilder
                '-- get general system and app information
                .Append(SysInfoToString())

                '-- get exception-specific information
                .Append("Exception Source:      ")
                Try
                    .Append(Ex.Source)
                Catch e As Exception
                    .Append(e.Message)
                End Try
                .Append(Environment.NewLine)

                .Append("Exception Type:        ")
                Try
                    .Append(Ex.GetType.FullName)
                Catch e As Exception
                    .Append(e.Message)
                End Try
                .Append(Environment.NewLine)

                .Append("Exception Message:     ")
                Try
                    .Append(Ex.Message)
                Catch e As Exception
                    .Append(e.Message)
                End Try
                .Append(Environment.NewLine)

                .Append("Exception Target Site: ")
                Try
                    .Append(Ex.TargetSite.Name)
                Catch e As Exception
                    .Append(e.Message)
                End Try
                .Append(Environment.NewLine)

                Try
                    Dim x As String = EnhancedStackTrace(Ex)
                    .Append(x)
                Catch e As Exception
                    .Append(e.Message)
                End Try
                .Append(Environment.NewLine)
            End With

            Return objStringBuilder.ToString

        Catch ex2 As Exception
            Return "Error '" & ex2.Message & "' while generating exception string"
        End Try
    End Function


    ''' <summary>
    ''' perform our string replacements for (app) and (contact), etc etc
    ''' also make sure More has default values if it is blank.
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <param name="WhatHappened"></param>
    ''' <param name="HowUserAffected"></param>
    ''' <param name="WhatUserCanDo"></param>
    ''' <param name="MoreDetails"></param>
    ''' <remarks></remarks>
    Private Sub ProcessStrings( _
            ByVal Ex As Exception, _
            ByRef WhatHappened As String, _
            ByRef HowUserAffected As String, _
            ByRef WhatUserCanDo As String, _
            ByRef MoreDetails As String)

        WhatHappened = ReplaceStringVals(Ex, WhatHappened)
        HowUserAffected = ReplaceStringVals(Ex, HowUserAffected)
        WhatUserCanDo = ReplaceStringVals(Ex, WhatUserCanDo)
        MoreDetails = ReplaceStringVals(Ex, GetDefaultMore(Ex, MoreDetails))
    End Sub


    ''' <summary>
    ''' make sure "More" text is populated with something useful
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <param name="MoreDetails"></param>
    ''' <returns></returns>
    Private Function GetDefaultMore(ByVal Ex As Exception, ByVal MoreDetails As String) As String
        Const DEFAULTMORE As String = "No further information is available."

        Dim objStringBuilder As New System.Text.StringBuilder
        With objStringBuilder
            If Len(MoreDetails) <> 0 Then
                .Append(MoreDetails)
            Else
                .Append(DEFAULTMORE)
            End If
            .Append(Environment.NewLine)
            .Append("If the problem persists, contact (contact).")
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)
            .Append("Basic technical information follows: " & Environment.NewLine)
            .Append("---" & Environment.NewLine)
            .Append(SysInfoToString(Ex))
        End With
        Return objStringBuilder.ToString
    End Function


    ''' <summary>
    ''' Retrieve the root assembly of the executing assembly
    ''' </summary>
    ''' <returns></returns>
    Private Function ParentAssembly() As System.Reflection.Assembly
        Static ParentAsm As System.Reflection.Assembly = Nothing
        If ParentAsm Is Nothing Then
            If System.Reflection.Assembly.GetEntryAssembly Is Nothing Then
                ParentAsm = System.Reflection.Assembly.GetCallingAssembly
            Else
                ParentAsm = System.Reflection.Assembly.GetEntryAssembly
            End If
        End If
        Return ParentAsm
    End Function


    ''' <summary>
    ''' gather some system information that is helpful to diagnosing
    ''' exception
    ''' </summary>
    ''' <param name="Ex"></param>
    ''' <returns></returns>
    Private Function SysInfoToString(Optional ByVal Ex As Exception = Nothing) As String
        Dim objStringBuilder As New System.Text.StringBuilder

        With objStringBuilder
            'If Err.Number <> 0 Then
            '    .Append("Error code:            ")
            '    .Append(Err.Number)
            '    .Append(Environment.NewLine)
            'End If

            'If Len(Err.Description) <> 0 Then
            '    .Append("Error Description:     ")
            '    .Append(Err.Description)
            '    .Append(Environment.NewLine)
            'End If

            ''---- report the line or ERL location as available
            'If Err.Line <> 0 Then
            '    .Append("Error Line:            ")
            '    .Append(Err.Line)
            '    If Err.Erl <> 0 AndAlso Err.Erl <> Err.Line Then
            '        .Append("  (Location " & Err.Erl.ToString & ")")
            '    End If
            '    .Append(Environment.NewLine)
            'ElseIf Err.Erl <> 0 Then
            '    .Append("Error Location:        ")
            '    .Append(Err.Erl)
            '    .Append(Environment.NewLine)
            'End If

            'If Err.Column <> 0 Then
            '    .Append("Error Column:          ")
            '    .Append(Err.Column)
            '    .Append(Environment.NewLine)
            'End If

            'If Len(Err.FileName) <> 0 Then
            '    .Append("Error Module:          ")
            '    .Append(Err.FileName)
            '    .Append(Environment.NewLine)
            'End If

            'If Len(Err.Method) <> 0 Then
            '    .Append("Error Method:          ")
            '    .Append(Err.Method)
            '    .Append(Environment.NewLine)
            'End If

            .Append("Date and Time:         ")
            .Append(DateTime.Now)
            .Append(Environment.NewLine)

            .Append("Machine Name:          ")
            Try
                .Append(Environment.MachineName)
            Catch e As Exception
                .Append(e.Message)
            End Try
            .Append(Environment.NewLine)

            .Append("IP Address:            ")
            .Append(GetCurrentIP())
            .Append(Environment.NewLine)

            .Append("Current User:          ")
            .Append(UserIdentity())
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)

            .Append("Application Domain:    ")
            Try
                .Append(System.AppDomain.CurrentDomain.FriendlyName())
            Catch e As Exception
                .Append(e.Message)
            End Try


            .Append(Environment.NewLine)
            .Append("Assembly Codebase:     ")
            Try
                .Append(ParentAssembly.CodeBase())
            Catch e As Exception
                .Append(e.Message)
            End Try
            .Append(Environment.NewLine)

            .Append("Assembly Full Name:    ")
            Try
                .Append(ParentAssembly.FullName)
            Catch e As Exception
                .Append(e.Message)
            End Try
            .Append(Environment.NewLine)

            .Append("Assembly Version:      ")
            Try
                .Append(ParentAssembly.GetName().Version().ToString)
            Catch e As Exception
                .Append(e.Message)
            End Try
            .Append(Environment.NewLine)

            .Append("Assembly Build Date:   ")
            Try
                .Append(AssemblyBuildDate(ParentAssembly).ToString)
            Catch e As Exception
                .Append(e.Message)
            End Try
            .Append(Environment.NewLine)
            .Append(Environment.NewLine)

            If Ex IsNot Nothing Then
                .Append(EnhancedStackTrace(Ex))
            End If

        End With

        Return objStringBuilder.ToString
    End Function


    ''' <summary>
    ''' turns a single stack frame object into an informative string
    ''' </summary>
    ''' <param name="FrameNum"></param>
    ''' <param name="sf"></param>
    ''' <returns></returns>
    Private Function StackFrameToString(ByVal FrameNum As Integer, ByVal sf As StackFrame) As String
        Dim sb As New System.Text.StringBuilder
        Dim intParam As Integer
        Dim mi As MemberInfo = sf.GetMethod

        With sb
            '-- build method name
            .Append("   ")
            Dim MethodName As String = mi.DeclaringType.Namespace & "." & mi.DeclaringType.Name & "." & mi.Name
            .Append(MethodName)
            'If FrameNum = 1 Then rCachedErr.Method = MethodName

            '-- build method params
            Dim objParameters() As ParameterInfo = sf.GetMethod.GetParameters()
            Dim objParameter As ParameterInfo
            .Append("(")
            intParam = 0
            For Each objParameter In objParameters
                intParam += 1
                If intParam > 1 Then .Append(", ")
                .Append(objParameter.Name)
                .Append(" As ")
                .Append(objParameter.ParameterType.Name)
            Next
            .Append(")")
            .Append(Environment.NewLine)

            '-- if source code is available, append location info
            .Append("       ")
            If InStr(1, Command, "/uselinemap", CompareMethod.Text) = 0 AndAlso _
                    Environ("uselinemap") = "" AndAlso _
                    (sf.GetFileName IsNot Nothing AndAlso sf.GetFileName.Length <> 0) Then
                '---- the PDB appears to be available, since the above elements are 
                '     not blank, so just use it's information

                .Append(System.IO.Path.GetFileName(sf.GetFileName))
                Dim Line = sf.GetFileLineNumber
                If Line <> 0 Then
                    .Append(": line ")
                    .Append(String.Format("{0:#0000}", Line))
                End If
                Dim col = sf.GetFileColumnNumber
                If col <> 0 Then
                    .Append(", col ")
                    .Append(String.Format("{0:#00}", sf.GetFileColumnNumber))
                End If
                '-- if IL is available, append IL location info
                If sf.GetILOffset <> StackFrame.OFFSET_UNKNOWN Then
                    .Append(", IL ")
                    .Append(String.Format("{0:#0000}", sf.GetILOffset))
                End If
            Else
                '---- the PDB is not available, so attempt to retrieve 
                '     any embedded linemap information
                Dim FileName As String = System.IO.Path.GetFileName(ParentAssembly.CodeBase)
                'If FrameNum = 1 Then rCachedErr.FileName = FileName
                .Append(FileName)
                '---- Get the native code offset and convert to a line number
                '     first, make sure our linemap is loaded
                Try
                    LineMap.AssemblyLineMaps.Add(sf.GetMethod.DeclaringType.Assembly)

                    Dim Line As Integer
                    Dim SourceFile As String = String.Empty
                    MapStackFrameToSourceLine(sf, Line, SourceFile)
                    If Line <> 0 Then
                        .Append(": Source File - ")
                        .Append(SourceFile)
                        .Append(": Line ")
                        .Append(String.Format("{0:#0000}", Line))
                    End If

                Catch ex As Exception
                    '---- just catch any exception here, if we can't load the linemap
                    '     oh well, we tried
                    OnWriteLog(ex, "Unable to load Line Map Resource")
                Finally
                    '-- native code offset is always available
                    Dim IL = sf.GetILOffset
                    If IL <> StackFrame.OFFSET_UNKNOWN Then
                        .Append(": IL ")
                        .Append(String.Format("{0:#00000}", IL))
                    End If
                End Try
            End If
            .Append(Environment.NewLine)
        End With
        Return sb.ToString
    End Function


    ''' <summary>
    ''' Map an address offset from a stack frame entry to a linenumber
    ''' using the Method name, the base address of the method and the
    ''' IL offset from the base address
    ''' </summary>
    ''' <param name="sf"></param>
    ''' <param name="Line"></param>
    ''' <param name="SourceFile"></param>
    ''' <remarks></remarks>
    Private Sub MapStackFrameToSourceLine(ByVal sf As StackFrame, ByRef Line As Integer, ByRef SourceFile As String)
        '---- first, get the base addr of the method
        '     if possible
        Line = 0
        SourceFile = String.Empty

        '---- you have to have symbols to do this
        If LineMap.AssemblyLineMaps.Count = 0 Then Exit Sub

        '---- first, check if for symbols for the assembly for this stack frame
        If Not LineMap.AssemblyLineMaps.Keys.Contains(sf.GetMethod.DeclaringType.Assembly.CodeBase) Then Exit Sub

        '---- retrieve the cache
        Dim alm = LineMap.AssemblyLineMaps(sf.GetMethod.DeclaringType.Assembly.CodeBase)

        '---- does the symbols list contain the metadata token for this method?
        Dim mi As MemberInfo = sf.GetMethod
        '---- Don't call this mdtoken or PostSharp will barf on it! Jeez
        Dim mdtokn As Long = mi.MetadataToken
        If Not alm.Symbols.ContainsKey(mdtokn) Then Exit Sub

        '---- all is good so get the line offset (as close as possible, considering any optimizations that
        '     might be in effect)
        Dim ILOffset = sf.GetILOffset
        If ILOffset <> StackFrame.OFFSET_UNKNOWN Then
            Dim Addr As Int64 = alm.Symbols(mdtokn).Address + ILOffset

            '---- now start hunting down the line number entry
            '     use a simple search. LINQ might make this easier
            '     but I'm not sure how. Also, a binary search would be faster
            '     but this isn't something that's really performance dependent
            Dim i As Integer = 1
            For i = alm.AddressToLineMap.Count - 1 To 0 Step -1
                If alm.AddressToLineMap(i).Address <= Addr Then
                    Exit For
                End If
            Next
            '---- since the address may end up between line numbers,
            '     always return the line num found
            '     even if it's not an exact match
            Line = alm.AddressToLineMap(i).Line
            SourceFile = alm.Names(alm.AddressToLineMap(i).SourceFileIndex)
        Else
            Exit Sub
        End If
    End Sub


    ''' <summary>
    ''' enhanced stack trace generator
    ''' </summary>
    ''' <param name="objStackTrace"></param>
    ''' <param name="SkipClassNameToSkip"></param>
    ''' <returns></returns>
    Private Function EnhancedStackTrace( _
            ByVal objStackTrace As StackTrace, _
            Optional ByVal SkipClassNameToSkip As String = "") As String

        Dim intFrame As Integer

        Dim sb As New System.Text.StringBuilder

        sb.Append(Environment.NewLine)
        sb.Append("---- Stack Trace ----")
        sb.Append(Environment.NewLine)

        Dim FrameNum As Integer = 0
        For intFrame = 0 To objStackTrace.FrameCount - 1
            Dim sf As StackFrame = objStackTrace.GetFrame(intFrame)

            If Len(SkipClassNameToSkip) > 0 AndAlso sf.GetMethod.DeclaringType.Name.IndexOf(SkipClassNameToSkip) > -1 Then
                '---- don't include frames with this name
                '     this lets of keep any ERR class frames out of
                '     the strack trace, they'd just be clutter
            Else
                FrameNum += 1
                sb.Append(StackFrameToString(FrameNum, sf))
            End If
        Next
        sb.Append(Environment.NewLine)

        Return sb.ToString
    End Function

    '--
    '-- enhanced stack trace generator (exception)
    '--
    Private Function InternalEnhancedStackTrace(ByVal objException As Exception) As String

        If objException Is Nothing Then
            Return EnhancedStackTrace(New StackTrace(True), ERRCLASSNAME)
        Else
            Return EnhancedStackTrace(New StackTrace(objException, True))
        End If
    End Function


    ''' <summary>
    ''' enhanced stack trace generator (no params)
    ''' </summary>
    ''' <returns></returns>
    Private Function EnhancedStackTrace() As String
        Return EnhancedStackTrace(New StackTrace(True), ERRCLASSNAME)
    End Function


    ''' <summary>
    ''' Experimental procedure to resolve overloaded method names
    ''' </summary>
    ''' <param name="mi"></param>
    ''' <returns></returns>
    Private Function DeriveFullMethodName(ByVal mi As MemberInfo) As String
        '---- I have to go through this because of overloaded methods
        '     If I have an overloaded method, there's basically no easy, direct way
        '     to correlate the SYMBOL from the PDB file that's now in the LINEMAP
        '     collections, with the MEMBERINFO from the stack frame.
        '
        '     Therefore, what I do is
        '     a) Since the symbols in the PDB are enumerated in linear order
        '        I number duplicates symbol, Symbol(2), Symbol(3), etc
        '     b) Once I have a stack frame, I get the membername from it, then
        '        retrieve all the members of the defining class, then
        '        find all those that match the name, then order them by 
        '        MetadataToken, which, fortunately, also appears to always
        '        be linear with how the methods are in the source, so that +SHOULD+
        '        line up with the naming from the PDB file.
        Dim c As Integer = 0
        Dim MethodName As String = ""
        Dim Members() As MemberInfo = mi.ReflectedType.GetMembers(BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static)

        Dim list As System.Collections.Generic.SortedList(Of Int64, MemberInfo) = New System.Collections.Generic.SortedList(Of Int64, MemberInfo)

        For Each mi2 As MemberInfo In Members
            If mi2.Name = mi.Name Then
                If Not list.ContainsKey(mi2.MetadataToken) Then
                    list.Add(mi2.MetadataToken, mi2)
                End If
            End If
        Next

        Dim Tokens As String = ""
        For i = 0 To list.Count - 1
            If list.Values(i).Name = mi.Name Then
                c += 1
                Tokens &= mi.Name & ":" & list.Values(i).MetadataToken.ToString & ": Order=" & i & vbCrLf
                If list.Values(i).MetadataToken = mi.MetadataToken Then
                    '---- this is the one we care about
                    MethodName = mi.DeclaringType.Namespace & "." & mi.DeclaringType.Name & "." & mi.Name
                    If c > 1 Then
                        '---- gotta append the instance count
                        MethodName &= "(" & c.ToString & ")"
                    End If
                    Exit For
                End If
            End If
        Next
        Return MethodName
    End Function


    ''' <summary>
    ''' retrieve user identity with fallback on error to safer method
    ''' </summary>
    ''' <returns></returns>
    Private Function UserIdentity() As String

        Dim strTemp As String = CurrentWindowsIdentity()
        If strTemp = "" Then
            strTemp = CurrentEnvironmentIdentity()
        End If
        Return strTemp
    End Function


    ''' <summary>
    '''  exception-safe WindowsIdentity.GetCurrent retrieval returns "domain\username"
    '''  per MS, this sometimes randomly fails with "Access Denied" particularly on NT4
    ''' </summary>
    ''' <returns></returns>
    Private Function CurrentWindowsIdentity() As String
        Try
            Return System.Security.Principal.WindowsIdentity.GetCurrent.Name()
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' exception-safe "domain\username" retrieval from Environment
    ''' </summary>
    ''' <returns></returns>
    Private Function CurrentEnvironmentIdentity() As String
        Try
            Return System.Environment.UserDomainName & "\" & System.Environment.UserName
        Catch ex As Exception
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' get IP address of this machine
    ''' not an ideal method for a number of reasons (guess why!)
    ''' but the alternatives are very ugly
    ''' </summary>
    ''' <returns></returns>
    Private Function GetCurrentIP() As String
        Try
            Return System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList(0).ToString
        Catch ex As Exception
            Return "127.0.0.1"
        End Try
    End Function


    ''' <summary>
    ''' Replace generic constants in a string with some built in values
    ''' </summary>
    ''' <param name="InputString"></param>
    ''' <returns></returns>
    Private Function ReplaceStringVals(ByVal Ex As Exception, ByVal InputString As String) As String
        If InputString Is Nothing Then
            InputString = ""
        Else
            GetAssemblyAttribs()
            InputString = InputString.Replace("(filename)", _AppFileName)
            InputString = InputString.Replace("(fullname)", _AppFullName)
            InputString = InputString.Replace("(company)", _AppCompany)
            InputString = InputString.Replace("(product)", _AppProduct)
            InputString = InputString.Replace("(app)", _AppProduct)
            InputString = InputString.Replace("(exception)", If(Ex Is Nothing, "", Ex.Message))
        End If
        Return InputString
    End Function


    ''' <summary>
    ''' If there is an entryassembly, retrieve it, otherwise, just retrieve
    ''' the calling assembly
    ''' </summary>
    ''' <returns></returns>
    Private Function GetEntryAssembly() As Reflection.Assembly
        If System.Reflection.Assembly.GetEntryAssembly Is Nothing Then
            Return System.Reflection.Assembly.GetCallingAssembly
        Else
            Return System.Reflection.Assembly.GetEntryAssembly
        End If
    End Function


    ''' <summary>
    ''' Retrieve the specific assembly attributes we care about
    ''' for error reporting purposes
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetAssemblyAttribs()
        '---- if we've already retrieved all this info
        '     there's no need to do it again
        If Len(_AppVersion) > 0 Then Exit Sub

        Dim objAssembly As System.Reflection.Assembly = GetEntryAssembly()
        Dim objAttributes() As Object = objAssembly.GetCustomAttributes(False)

        For Each objAttribute As Object In objAttributes
            Select Case objAttribute.GetType().ToString()
                Case "System.Reflection.AssemblyProductAttribute"
                    _AppProduct = CType(objAttribute, AssemblyProductAttribute).Product.ToString
                Case "System.Reflection.AssemblyCompanyAttribute"
                    _AppCompany = CType(objAttribute, AssemblyCompanyAttribute).Company.ToString
                Case "System.Reflection.AssemblyTitleAttribute"
                    _AppTitle = CType(objAttribute, AssemblyTitleAttribute).Title.ToString
                Case "System.Reflection.AssemblyDescriptionAttribute"
                    _AppDescription = CType(objAttribute, AssemblyDescriptionAttribute).Description.ToString
                Case Else
            End Select
        Next

        '-- add some extra values that are not in the AssemblyInfo, but nice to have
        _AppCodeBase = objAssembly.CodeBase.Replace("file:///", "")
        _AppBuildDate = AssemblyBuildDate(objAssembly).ToString
        _AppVersion = objAssembly.GetName.Version.ToString
        _AppFullName = objAssembly.FullName

        '-- we must have certain assembly keys to proceed.
        If Len(_AppProduct) = 0 Then
            '---- Just use the filename
            _AppProduct = _AppFileName
        End If
        If Len(_AppCompany) = 0 Then
            '---- just use the filename
            _AppCompany = _AppFileName
        End If
    End Sub


    ''' <summary>
    ''' returns build datetime of assembly
    ''' assumes default assembly value in AssemblyInfo:
    ''' {Assembly: AssemblyVersion("1.0.*")}
    '''
    ''' filesystem create time is used, if revision and build were overridden by user
    ''' </summary>
    ''' <param name="objAssembly"></param>
    ''' <param name="bForceFileDate"></param>
    ''' <returns></returns>
    Private Function AssemblyBuildDate(ByVal objAssembly As System.Reflection.Assembly, _
                                       Optional ByVal bForceFileDate As Boolean = False) As DateTime
        Dim objVersion As System.Version = objAssembly.GetName.Version
        Dim dtBuild As DateTime

        If bForceFileDate Then
            dtBuild = AssemblyFileTime(objAssembly)
        Else
            dtBuild = CType("01/01/2000", DateTime). _
                AddDays(objVersion.Build). _
                AddSeconds(objVersion.Revision * 2)
            If TimeZone.IsDaylightSavingTime(DateTime.Now, TimeZone.CurrentTimeZone.GetDaylightChanges(DateTime.Now.Year)) Then
                dtBuild = dtBuild.AddHours(1)
            End If
            If dtBuild > DateTime.Now Or objVersion.Build < 730 Or objVersion.Revision = 0 Then
                dtBuild = AssemblyFileTime(objAssembly)
            End If
        End If

        Return dtBuild
    End Function


    ''' <summary>
    ''' exception-safe file attrib retrieval; we don't care if this fails
    ''' </summary>
    ''' <param name="objAssembly"></param>
    ''' <returns></returns>
    Private Function AssemblyFileTime(ByVal objAssembly As System.Reflection.Assembly) As DateTime
        Try
            Return System.IO.File.GetLastWriteTime(objAssembly.Location)
        Catch ex As Exception
            Return DateTime.MinValue
        End Try
    End Function


    ''' <summary>
    ''' Relies on an Extension method that must be externally supplied
    ''' </summary>
    ''' <param name="Msg"></param>
    ''' <remarks></remarks>
    Private Sub OnWriteLog(ByVal Exception As Exception, ByRef Msg As String)
        Exception.Log(Msg)
    End Sub
End Module
#End Region

