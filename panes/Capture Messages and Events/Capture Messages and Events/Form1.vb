Imports System.Runtime.InteropServices
Public Class Form1

    Dim App As femap.model
    Dim AppLoaded As Boolean = False    'Global Variable True once the Pane has been registered and connected

    Private Const WM_MOUSEWHEEL As Integer = &H20A
    Private Const WM_COPYDATA As Integer = &H4A

    Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

    Private Structure COPYDATASTRUCT
        Public dwData As IntPtr
        Public cbData As Integer
        Public lpData As String
    End Structure


    Public Sub CaptureMessages()
        Dim rc As Long
        App = GetObject(, "femap.model")

        rc = App.feAppRegisterAddInPane(True, Me.Handle, Me.Handle, True, False, 2, 0)
        rc = App.feAppRegisterMessageHandler(True, Me.Handle)
        If rc = femap.zReturnCode.FE_OK Then
            AppLoaded = True
        End If
        'WM_FEMAP_MESSAGE = RegisterWindowMessage("FE_EVENT_MESSAGE")
        'PreviousWindowProc = SetWindowLong(Me.Handle, GWL_WNDPROC, AddressOf MsgWindowProc)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_MOUSEWHEEL
                ' ...do something...
                Me.FemapMsgHandlerListBox.Items.Add("scrolling on Form")
                'Exit Sub ' Suppress Default Action (because we don't reach the last line below
            Case WM_COPYDATA
                'Me.FemapMsgHandlerListBox.Items.Add(m.Msg)
                Dim cds As COPYDATASTRUCT = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(COPYDATASTRUCT)), COPYDATASTRUCT)
                Me.FemapMsgHandlerListBox.Items.Add(cds.lpData) 'Print Messages

        End Select

        MyBase.WndProc(m)
    End Sub
    '================================================================
    ' Stop capturing messages
    '----------------------------------------------------------------
    Public Sub ReleaseMessages()
        Dim rc As Long
        'If PreviousWindowProc Then
        rc = App.feAppRegisterAddInPane(False, Me.Handle, Me.Handle, True, False, 2, 0)
        rc = App.feAppRegisterMessageHandler(False, Me.Handle)
        Environment.Exit(0)
        'rc = SetWindowLong(Me.Handle, GWL_WNDPROC, PreviousWindowProc)
        'PreviousWindowProc = 0
        'End If
    End Sub

    '------------------------- Form events ---------------------------------------
    Private Sub CaptureMsgButton_Click(sender As Object, e As EventArgs) Handles CaptureMsgButton.Click
        CaptureMessages()
    End Sub

    Private Sub ListModelInfoButton_Click(sender As Object, e As EventArgs) Handles ListModelInfoButton.Click
        Dim rc As Long
        rc = App.feFileProgramRun(False, True, False, "{LI}")
    End Sub

    Private Sub ReleaseMsgButton_Click(sender As Object, e As EventArgs) Handles ReleaseMsgButton.Click
        ReleaseMessages()
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        ReleaseMessages()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        If AppLoaded Then
            App.feAppRegisterAddInPane(False, Me.Handle, Me.Handle, True, True, 2, 0) 'Disconnect our pane
        End If
        Close() 'Exit
    End Sub
End Class