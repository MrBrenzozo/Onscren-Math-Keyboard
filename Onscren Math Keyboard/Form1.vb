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
        KeyboardButton(130, 70, 0, 375, "(x-y)(x+y)", "^" + "6") 'binomial rule'
        KeyboardButton(70, 70, 0, 505, "( )", "%" + "8")
        KeyboardButton(70, 70, 0, 575, "{ }", "%" + "c")
        KeyboardButton(70, 70, 0, 645, ChrW(&H3C), "%" + "a") 'less than
        KeyboardButton(70, 70, 0, 715, ChrW(&H3E), "%" + "o") 'greater than
        KeyboardButton(70, 70, 0, 785, ChrW(&H2260), "%" + "e") 'not equal
        KeyboardButton(70, 70, 0, 855, ChrW(&H2155), "%" + "9") 'reciprocal
        KeyboardButton(70, 70, 0, 925, ChrW(&H221A), "%" + "6") 'radical
        KeyboardButton(70, 70, 0, 995, "x" + ChrW(&H207F), "%" + "7") 'exponent
        KeyboardButton(80, 70, 0, 1065, "Sin()", "^" + "%" + "0") ' sin
        KeyboardButton(80, 70, 0, 1145, "Cos()", "^" + "%" + "q") ' cos
        KeyboardButton(80, 70, 0, 1225, "Tan()", "^" + "%" + "w") ' tan
        KeyboardButton(55, 70, 0, 1305, ChrW(&HB0), "^" + "%" + "4") 'degree
        KeyboardButton(55, 70, 0, 1360, ChrW(&H3C0), "^" + "%" + "6") 'pi
        KeyboardButton(55, 70, 0, 1415, ChrW(&H3C9), "^" + "%" + "x") 'omega
        KeyboardButton(70, 70, 0, 1470, ChrW(&H2192), "^" + "%" + "9") 'arrow
        'KeyboardButton(70, 70, 0, 1540, ChrW(&H2211), "^" + "9") 'sigma
        'KeyboardButton(70, 70, 0, 1610, ChrW(&H222B), "^" + "t") 'integral
        'KeyboardButton(70, 70, 0, 1680, " ", "%" + "+" + "5") 'unused




        ' Row Two
        KeyboardButton(95, 70, 50, 0, "Pythag" & vbCrLf & "Theroem", "^" + "4")
        KeyboardButton(105, 70, 50, 95, "Quadratic" & vbCrLf & "Equation", "^" + "3")
        KeyboardButton(90, 70, 50, 200, "log" + ChrW(&H2090) + "(x)=n", "%" + "v") 'Logarithm
        KeyboardButton(85, 70, 50, 290, ChrW(&H192) + "(x)={", "^" + "8") 'f(x)


        KeyboardButton(130, 70, 70, 375, "x" + ChrW(&HB2) + " + 2xy + y" + ChrW(&HB2), "^" + "7") 'other binomial rule
        KeyboardButton(70, 70, 70, 505, "| |", "%" + "i")
        KeyboardButton(70, 70, 70, 575, "[ ]", "%" + "x")
        KeyboardButton(70, 70, 70, 645, ChrW(&H2264), "%" + "k") 'less than or equal too
        KeyboardButton(70, 70, 70, 715, ChrW(&H2265), "%" + "d") 'greater than or equal too
        KeyboardButton(70, 70, 70, 785, ChrW(&H2248), "%" + "t") ' approximate
        KeyboardButton(70, 70, 70, 855, ChrW(&H215C), "%" + "0") 'fraction
        KeyboardButton(70, 70, 70, 925, ChrW(&H2213), "%" + "u") 'plus or minus
        KeyboardButton(70, 70, 70, 995, "X" + ChrW(&H2099), "%" + "b") 'subscript
        KeyboardButton(80, 70, 70, 1065, "Sec()", "^" + "%" + "h")
        KeyboardButton(80, 70, 70, 1145, "Csc()", "^" + "%" + "g")
        KeyboardButton(80, 70, 70, 1225, "Cot()", "^" + "%" + "a")
        KeyboardButton(55, 70, 70, 1305, ChrW(&H2220), "^" + "%" + "8") ' Angle
        KeyboardButton(55, 70, 70, 1360, ChrW(&H3B8), "^" + "%" + "j") 'theta
        KeyboardButton(55, 70, 70, 1415, ChrW(&H3BC), "^" + "%" + "5") 'micro
        KeyboardButton(70, 70, 70, 1470, ChrW(&H221E), "^" + "%" + "7") 'infinity
        'KeyboardButton(70, 70, 70, 1540, ChrW(&H394), "%" + "+" + "t") 'delta
        'KeyboardButton(70, 70, 70, 1610, " ", "^" + "%" + "7") 'unused
        'KeyboardButton(70, 70, 70, 1680, " ", "^" + "%" + "7") 'unued





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