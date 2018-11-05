Public Class frmMain

    ' API Methods

    Public Declare Function IsWindowVisible Lib "user32.dll" _
Alias "IsWindowVisible" (ByVal hwnd As Integer) As Boolean

    Public Declare Function GetWindow Lib "user32.dll" _
Alias "GetWindow" (ByVal hwnd As Integer,
                   ByVal wCmd As Integer) As Integer

    Public Declare Function GetWindowLong Lib "user32.dll" _
Alias "GetWindowLongA" (ByVal hwnd As Integer,
                        ByVal nIndex As Integer) As Integer

    Public Declare Function GetParent Lib "user32.dll" _
Alias "GetParent" (ByVal hwnd As Integer) As Integer


    Public Declare Function SetForegroundWindow Lib "user32.dll" _
Alias "SetForegroundWindow" (ByVal hwnd As Integer) As Integer


    Private CapsLock As New CheckBox
    Private Key, Special As String
    Private Windows As New ArrayList
    Private Window As IntPtr


    ' Event Handlers

    Private Sub Key_Special(ByVal Sender As Button, ByVal e As System.EventArgs)
        If Special = "" Then
            Special = Sender.Tag
        Else
            Special = ""
        End If
    End Sub

    Private Sub Key_Click(ByVal Sender As Button,
                        ByVal e As System.EventArgs)
        Key = Sender.Tag
        If Key = "{SPACE}" Then Key = " " 'Convert {SPACE} to Space
        If Window <> 0 Then
            SetForegroundWindow(Window)
            SendKeys.SendWait(Special & IIf(CapsLock.Checked And Not Special <> "^", UCase(Key), Key))
            SetForegroundWindow(Window)
        End If
    End Sub

    ' Private Methods

    Private Function IsActiveWindow(ByVal hWnd As Integer) As Boolean
        Dim IsOwned As Boolean
        Dim Style As Integer
        IsOwned = GetWindow(hWnd, 4) <> 0
        Style = GetWindowLong(hWnd, -20)
        If Not IsWindowVisible(hWnd) Then Return False ' Not Visible
        If GetParent(hWnd) <> 0 Then Return False ' Has Parent
        If (Style And &H80) <> 0 And Not IsOwned Then Return False ' Is Tooltip
        If (Style And &H40000) = 0 And IsOwned Then Return False ' Has Owner
        If Process.GetCurrentProcess.MainWindowHandle = hWnd Then Return False
        Return True ' Window Valid
    End Function

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        ' Row One
        KeyboardButton(100, 70, 0, 350, "(x-y)(x+y)", "^" + "6") 'binomial rule
        KeyboardButton(70, 70, 0, 450, "( )", "%" + "8")
        KeyboardButton(70, 70, 0, 520, "{ }", "%" + "c")
        KeyboardButton(70, 70, 0, 590, ChrW(&H3C), "%" + "a") 'less than
        KeyboardButton(70, 70, 0, 660, ChrW(&H3E), "%" + "o") 'greater than
        KeyboardButton(70, 70, 0, 730, ChrW(&H2260), "%" + "e") 'not equal
        KeyboardButton(70, 70, 0, 800, "1 " + ChrW(&H2044) + " x", "%" + "9") 'reciprocal
        KeyboardButton(70, 70, 0, 870, ChrW(&H221A), "%" + "6") 'radical
        KeyboardButton(70, 70, 0, 940, "x" + ChrW(&H207F), "%" + "7") 'exponent
        KeyboardButton(70, 70, 0, 1010, "Sin()", "^" + "%" + "0") ' sin
        KeyboardButton(70, 70, 0, 1080, "Cos()", "^" + "%" + "q") ' cos
        KeyboardButton(70, 70, 0, 1150, "Tan()", "^" + "%" + "w") ' tan
        KeyboardButton(55, 70, 0, 1220, ChrW(&HB0), "^" + "%" + "4") 'degree
        KeyboardButton(55, 70, 0, 1275, ChrW(&H3C0), "^" + "%" + "6") 'pi
        KeyboardButton(55, 70, 0, 1330, ChrW(&H3C9), "^" + "%" + "x") 'omega
        KeyboardButton(70, 70, 0, 1385, ChrW(&H2192), "^" + "%" + "9") 'arrow



        ' Row Two
        KeyboardButton(95, 70, 50, 0, "Pythag" & vbCrLf & "Theroem", "^" + "4")
        KeyboardButton(95, 70, 50, 95, "Quadratic" & vbCrLf & "Equation", "^" + "3")
        KeyboardButton(90, 70, 50, 190, "log" + ChrW(&H2090) + "(x)=n", "%" + "v") 'Logarithm
        KeyboardButton(70, 70, 50, 280, ChrW(&H192) + "(x)=", "^" + "8") 'f(x)


        KeyboardButton(100, 70, 70, 350, "x" + ChrW(&HB2) + " + 2xy + y" + ChrW(&HB2), "^" + "7") 'other binomial rule
        KeyboardButton(70, 70, 70, 450, "| |", "%" + "i")
        KeyboardButton(70, 70, 70, 520, "[ ]", "%" + "x")
        KeyboardButton(70, 70, 70, 590, ChrW(&H2264), "%" + "k") 'less than or equal too
        KeyboardButton(70, 70, 70, 660, ChrW(&H2265), "%" + "d") 'greater than or equal too
        KeyboardButton(70, 70, 70, 730, ChrW(&H2248), "%" + "t") ' approximate
        KeyboardButton(70, 70, 70, 800, "  x " + ChrW(&H2044) + "  y", "%" + "0") 'fraction
        KeyboardButton(70, 70, 70, 870, ChrW(&H2213), "%" + "u") 'plus or minus
        KeyboardButton(70, 70, 70, 940, "x" + ChrW(&H2099), "%" + "b") 'subscript
        KeyboardButton(70, 70, 70, 1010, "Sec()", "^" + "%" + "h")
        KeyboardButton(70, 70, 70, 1080, "Csc()", "^" + "%" + "g")
        KeyboardButton(70, 70, 70, 1150, "Cot()", "^" + "%" + "a")
        KeyboardButton(55, 70, 70, 1220, ChrW(&H2220), "^" + "%" + "8") ' Angle
        KeyboardButton(55, 70, 70, 1275, ChrW(&H3B8), "^" + "%" + "j") 'theta
        KeyboardButton(55, 70, 70, 1330, ChrW(&H3BC), "^" + "%" + "5") 'micro
        KeyboardButton(70, 70, 70, 1385, ChrW(&H221E), "^" + "%" + "7") 'infinity






    End Sub

    Private Sub cboWindows_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboWindows.DropDown

        Windows.Clear()
        cboWindows.Items.Clear()
        For Each Item As Process In Process.GetProcesses
            If IsActiveWindow(Item.MainWindowHandle) _
            And Item.MainWindowTitle <> "" Then
                Windows.Add(Item.MainWindowHandle)
                cboWindows.Items.Add(Item.MainWindowTitle)
            End If
        Next

    End Sub

    Private Sub cboWindows_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboWindows.SelectedIndexChanged

        If cboWindows.SelectedItem <> Nothing Then
            Window = Windows.Item(cboWindows.SelectedIndex)
        End If

    End Sub

    Private Sub KeyboardButton(ByVal Width As Integer, ByVal Height As Integer,
                             ByVal Top As Integer, ByVal Left As Integer,
                             Optional ByVal Text As String = "",
                             Optional ByVal Tag As String = "",
                             Optional ByVal Special As Boolean = False)
        Dim Button As New Button
        Button.Size = New Size(Width, Height)
        Button.Location = New Point(Left, Top)
        Button.Text = Text
        Button.Tag = Tag
        If Special Then
            AddHandler Button.Click, AddressOf Key_Special
        Else
            AddHandler Button.Click, AddressOf Key_Click
        End If
        Controls.Add(Button)
    End Sub

End Class