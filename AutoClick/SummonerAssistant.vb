Imports System.Runtime.InteropServices
Public Class SummonerAssistant

#Region "Findwindow"
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(
     ByVal lpClassName As String,
     ByVal lpWindowName As String) As IntPtr
    End Function
    'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Private Declare Auto Function GetWindowText Lib "user32" (ByVal hWnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
#End Region

#Region "GetDC"
    Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Integer) As Integer
    Declare Function GetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Integer) As Integer
#End Region

#Region "Mouse"
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Integer
    Declare Function SetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal x As Integer, ByVal y As Integer) As Integer '設定滑鼠游標
    Const MOUSEEVENTF_LEFTDOWN = &H2
    Const MOUSEEVENTF_LEFTUP = &H4
    Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Const MOUSEEVENTF_MIDDLEUP = &H40
    Const MOUSEEVENTF_MOVE = &H1
    Const MOUSEEVENTF_ABSOLUTE = &H8000
    Const MOUSEEVENTF_RIGHTDOWN = &H8
    Const MOUSEEVENTF_RIGHTUP = &H10
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
#End Region

    Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
    Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer

    Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    Dim rectwin As RECT
    Dim TargetWindowRect As RECT

#Region "CaptureDC"
    Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
    Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Integer) As Integer
    Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
    Private Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
    Private Declare Function BitBlt Lib "GDI32" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Integer) As Integer
    Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Integer) As Integer
    Private Declare Function DeleteObject Lib "GDI32" (ByVal hObj As Integer) As Integer
    Const SRCCOPY As Integer = &HCC0020

#End Region

#Region "SetParent"
    Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Int32, ByVal hWndNewParent As Int32) As Boolean
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
#End Region

#Region "Message"
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Const WM_MOUSE_MOVE = &H200
    Public Const WM_LBUTTON_DOWN = &H201 '0x C# 16x = &H
    Public Const WM_LBUTTON_UP = &H202
    Public Const WM_KEYDOWN = &H100
    Public Const WM_KEYUP = &H101
    Public Const WM_CHAR = &H102

    Public Function MAKELPARAM(ByVal l As Integer, ByVal h As Integer) As Integer
        Dim r As Integer = l + (h << 16)
        Return r
    End Function
#End Region

    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Integer, lpdwProcessId As Integer) As Integer
#Region "Basic"
    Dim hwnd, hwnd2, hwnd3 As Integer
    Dim bmpBackground, SubBackground As Bitmap
    Dim intWidth, intHeight As Integer
    Dim hsdc, hmdc As Integer
    Dim bmpHandle, OLDbmpHandle As Integer
    Dim releaseDC As Integer
    Dim index As Integer = 10
    Dim img(index) As Bitmap
    Dim pt(index) As Point
    Dim turn, sturn As Point
    Dim target As String = "BlueStacks Android PluginAndroid"
    Dim IndexCounter As Integer = 0
    Dim TimerIndexCheck As Integer = -1
    Dim SendClickTracker(10) As Integer

#End Region

    Function GetBitmap(Optional ByVal hwnd As Integer = 0, Optional ByVal BitWidth As Integer = -1, Optional ByVal BitHeight As Integer = -1) As Bitmap
        '13369376
        '&HCC0020

        Dim SRCCOPY As Integer = &HCC0020
        Dim SRCPAINT As Integer = &HEE0086
        Dim SRCAND As Integer = &H8800C6
        Dim SRCINVERT As Integer = &H660046
        Dim SRCERASE As Integer = &H440328
        Dim NOTSRCCOPY As Integer = &H330008
        Dim NOTSRCERASE As Integer = &H1100A6
        Dim MERGECOPY As Integer = &HC000CA
        Dim MERGEPAINT As Integer = &HBB0226
        Dim PATCOPY As Integer = &HF00021
        Dim PATPAINT As Integer = &HFB0A09
        Dim PATINVERT As Integer = &H5A0049
        Dim DSTINVERT As Integer = &H550009
        Dim BLACKNESS As Integer = &H42
        Dim WHITENESS As Integer = &HFF0062
        Dim CAPTUREBLT As Integer = &H40000000

        If BitWidth < 0 Then BitWidth = My.Computer.Screen.Bounds.Width
        If BitHeight < 0 Then BitHeight = My.Computer.Screen.Bounds.Height
        Dim Bhandle, DestDC, SourceDC As IntPtr
        'SourceDC = CreateDC("DISPLAY", Nothing, Nothing, 0)
        SourceDC = GetDC(hwnd)
        DestDC = CreateCompatibleDC(SourceDC)

        Bhandle = CreateCompatibleBitmap(SourceDC, BitWidth, BitHeight)
        SelectObject(DestDC, Bhandle)
        BitBlt(DestDC, 0, 0, BitWidth, BitHeight, SourceDC, 0, 0, SRCCOPY)

        Dim bmp As Bitmap = Image.FromHbitmap(Bhandle)

        'releaseDC(hwnd, SourceDC)
        DeleteDC(hwnd)
        DeleteDC(SourceDC)
        DeleteDC(DestDC)
        DeleteObject(Bhandle)
        GC.Collect()

        Return bmp

    End Function

    Public Sub CaptureScreen()
        hsdc = GetWindowDC(Me.Handle)
        'hsdc = CreateDC("DISPLAY", "", "", "")
        hmdc = CreateCompatibleDC(hsdc)

        'intWidth = GetDeviceCaps(hsdc, 8) '8 or 10
        'intHeight = GetDeviceCaps(hsdc, 10)
        intWidth = rectwin.Right - rectwin.Left
        intHeight = rectwin.Bottom - rectwin.Top

        bmpHandle = CreateCompatibleBitmap(hsdc, intWidth, intHeight)

        OLDbmpHandle = SelectObject(hmdc, bmpHandle)
        releaseDC = BitBlt(hmdc, 0, 0, intWidth, intHeight, hsdc, 0, 0, 13369376)
        bmpHandle = SelectObject(hmdc, OLDbmpHandle)

        releaseDC = DeleteDC(hsdc)
        releaseDC = DeleteDC(hmdc)

        bmpBackground = Image.FromHbitmap(New IntPtr(bmpHandle))
        DeleteObject(bmpHandle)
    End Sub

    Public Sub CaptureSubScreen()
        hsdc = GetWindowDC(Me.Handle)
        'hsdc = CreateDC("DISPLAY", "", "", "")
        hmdc = CreateCompatibleDC(hsdc)

        intWidth = GetDeviceCaps(hsdc, 8) '8 or 10
        intHeight = GetDeviceCaps(hsdc, 10)
        bmpHandle = CreateCompatibleBitmap(hsdc, intWidth, intHeight)

        OLDbmpHandle = SelectObject(hmdc, bmpHandle)
        releaseDC = BitBlt(hmdc, 0, 0, intWidth, intHeight, hsdc, 0, 0, 13369376)
        bmpHandle = SelectObject(hmdc, OLDbmpHandle)

        'releaseDC = DeleteDC(hsdc)
        releaseDC = DeleteDC(hmdc)

        SubBackground = Image.FromHbitmap(New IntPtr(bmpHandle))
        DeleteObject(bmpHandle)
    End Sub

    Public Function SearchBitmap(mainBmp As Bitmap, childBmp As Bitmap, Optional ByVal LocationX As Integer = -1, Optional ByVal LocationY As Integer = -1) As Point
        'mainBmp 大圖 childBmp 目標圖 locationX Y = 從哪裡找 (效率)
        Try
            Dim c0 As Color = childBmp.GetPixel(0, 0)
            Dim rp As Point = New Point(-1, -1)
            If LocationX = -1 And LocationY = -1 Then
                Dim endsearch As Boolean = False
                For i = 0 To mainBmp.Width - 1
                    For j = 0 To mainBmp.Height - 1
                        If i + childBmp.Width < mainBmp.Width And j + childBmp.Height < mainBmp.Height Then
                            If mainBmp.GetPixel(i, j) = c0 Then
                                Dim isequal As Boolean = True
                                For i2 = 1 To childBmp.Width - 1
                                    For j2 = 1 To childBmp.Height - 1
                                        If mainBmp.GetPixel(i + i2, j + j2) <> childBmp.GetPixel(i2, j2) Then
                                            isequal = False
                                            Exit For
                                        End If
                                    Next
                                    If isequal = False Then
                                        Exit For
                                    End If
                                Next
                                If isequal Then
                                    rp = New Point(i, j)
                                    endsearch = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If endsearch Then
                        Exit For
                    End If
                Next
            Else
                If LocationX + childBmp.Width < mainBmp.Width And LocationY + childBmp.Height < mainBmp.Height Then
                    Dim isequal As Boolean = True
                    For i = 0 To childBmp.Width - 1
                        For j = 0 To childBmp.Height - 1
                            If mainBmp.GetPixel(i + LocationX, j + LocationY) <> childBmp.GetPixel(i, j) Then
                                isequal = False
                                Exit For
                            End If
                        Next
                        If isequal = False Then
                            Exit For
                        End If
                    Next
                    If isequal Then
                        rp = New Point(LocationX, LocationY)
                    End If
                End If
            End If
            Return rp
        Catch ex As Exception
            Return New Point(-1, -1)
        End Try
    End Function

    Private Sub AssistantForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        index = 0
        'img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\WaterGuardTower\\Crystal.png") : pt(index) = New Point(1168, 445) : index += 1
        'img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\WaterGuardTower\\AgainButton.png") : pt(index) = New Point(372, 458) : index += 1
        'img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\WaterGuardTower\\ConfirmButton.png") : pt(index) = New Point(784, 737) : index += 1
        'img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\WaterGuardTower\\CancelButton.png") : pt(index) = New Point(1032, 307) : index += 1

        img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\FireMountain\\Crystal.png") : pt(index) = New Point(1168, 445) : index += 1
        img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\FireMountain\\SellButton.png") : pt(index) = New Point(676, 671) : index += 1
        img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\FireMountain\\CheckButton.png") : pt(index) = New Point(-1, -1) : index += 1
        img(index) = Image.FromFile(System.Environment.CurrentDirectory + "\\FireMountain\\AgainButton.png") : pt(index) = New Point(375, 457) : index += 1

        index = 0
        hwnd = FindWindow(vbNullString, "BlueStacks")

        GetWindowRect(hwnd, rectwin)
        Console.WriteLine(rectwin.Left.ToString + " " + rectwin.Right.ToString + " " + rectwin.Top.ToString + " " + rectwin.Bottom.ToString)

        Me.Text = rectwin.Left.ToString + " " + rectwin.Right.ToString + " " + rectwin.Top.ToString + " " + rectwin.Bottom.ToString
        Me.Width = rectwin.Right - rectwin.Left
        Me.Height = rectwin.Bottom - rectwin.Top

        If (hwnd) Then
            EnumWindows(hwnd)
        End If
        Me.Text = hwnd2.ToString
        hwnd2 = FindWindowEx(hwnd, IntPtr.Zero, vbNullString, vbNullString)
        SetParent(hwnd, PictureBox1.Handle)
        SetWindowPos(hwnd, 0, 0, 0, Me.Width, Me.Height, 4)

        Me.Text = "hwnd = " + hwnd.ToString() + " , hwnd2 = " + hwnd2.ToString()
        Assistance_Timer.Enabled = False
        KeyboardTrigger_Timer.Enabled = True
        Assistance_Timer.Interval = 100
        KeyboardTrigger_Timer.Interval = 100
    End Sub

    Private Sub AssistantForm_Closing(sender As Object, e As EventArgs) Handles MyBase.FormClosing
        If Not (hwnd = IntPtr.Zero) Then
            Dim p() As Process = Process.GetProcessesByName("Bluestacks")

            p(0).Kill()

            p = Process.GetProcessesByName("HD-Agent")
            p(0).Kill()

            p = Process.GetProcessesByName("HD-Player")
            p(0).Kill()

            p = Process.GetProcessesByName("BstkSVC")
            p(0).Kill()

            hwnd = IntPtr.Zero
        End If
    End Sub

    Sub EnumWindows(h As Integer)
        Dim h2 As Integer = 0
        Dim lpString As String = New String(Chr(0), 255)
        Dim Ret As Integer '= GetWindowText(Me.Handle, lpString, 255)
        Console.WriteLine("h : " + h.ToString())
        While True
            h2 = FindWindowEx(h, h2, vbNullString, vbNullString)
            Console.WriteLine("h2 : " + h2.ToString())
            If h2 Then
                Ret = GetWindowText(h2, lpString, 255)
                If String.Compare(lpString, target, False) = 0 Then
                    hwnd2 = h2
                    Exit While
                End If
                EnumWindows(h2)
            Else
                Exit While
            End If
        End While
    End Sub

    Function SubBG(bindex As Integer) As Boolean
        Try
            Dim rp As Point = SearchBitmap(SubBackground, img(bindex), pt(bindex).X, pt(bindex).Y)
            If rp.X = -1 And rp.Y = -1 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub SendClick(index As Integer)
        Dim key As Integer
        If (index = 0) Then
            key = Keys.A
            Console.WriteLine("$Click A")
        ElseIf index = 1 Then
            key = Keys.B
            Console.WriteLine("$Click B")
        ElseIf index = 2 Then
            key = Keys.C
            Console.WriteLine("$Click C")
        ElseIf index = 3 Then
            key = Keys.D
            Console.WriteLine("$Click D")
        ElseIf index = 4 Then
            key = Keys.E
            Console.WriteLine("$Click E")
        ElseIf index = 5 Then
            key = Keys.F
            Console.WriteLine("$Click F")
        End If
        SendClickTracker(IndexCounter) = index
        IndexCounter += 1
        PostMessage(hwnd2, WM_KEYDOWN, key, MAKELPARAM(key, WM_KEYDOWN))
        PostMessage(hwnd2, WM_KEYUP, key, MAKELPARAM(key, WM_KEYUP))
        Sleep(500)

    End Sub

    Private Sub SimulatePressPtrScr()

        'Simulate Press PrtSrc Button

        SendKeys.Send("%")
        System.Windows.Forms.SendKeys.Send("%{PRTSC}")
        Sleep(100)
        Dim bmp As Bitmap = CType(Clipboard.GetDataObject().GetData(DataFormats.Bitmap), Bitmap)
        'bmp.Save("C:\Users\ReplacedToy\Desktop\test1.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
        Me.BackgroundImage = bmpBackground
        bmp.Save("test1.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
        bmp.Dispose()
        Me.BackgroundImage = Nothing
        Me.Refresh()
        SearchIndexImage.Enabled = False

    End Sub

    Private Sub Assistance_Timer_Tick(sender As System.Object, e As System.EventArgs) Handles Assistance_Timer.Tick
        CaptureScreen()
        PictureBox1.Image = bmpBackground

        turn = SearchBitmap(bmpBackground, img(index), pt(index).X, pt(index).Y)
        If (turn <> pt(index)) Then
            Me.Text = "AutoClick Searching ... Index : " + index.ToString() + " (False)"
        Else
            Me.Text = "AutoClick Searching ... Index : " + index.ToString() + " (True) , SearchPoint = (" + turn.X.ToString() + " , " + turn.Y.ToString() + ")"

            If (index = 0) Then
                SendClick(index)
                Sleep(200)
                SendClick(index)
                Sleep(400)
                SendClick(index)
                Sleep(800)

                SendClick(1)
                Sleep(500)
                SendClick(2)
                Sleep(500)

                index = 4
                'CaptureSubScreen()
                'For st = 1 To 3
                '    If (SubBG(st) = True) Then
                '        index = st
                '        Exit For
                '    End If
                'Next
                Sleep(100)
                SendClickTracker_Track()
                'SubBackground.Dispose()
            ElseIf index = 1 Then
                SendClick(index) ' 1 = Press B
                index = 4
            ElseIf index = 2 Or index = 3 Then
                SendClick(2) ' 2 = Press C
                index = 4
            ElseIf index = 4 Then
                SendClick(3) ' = Press D
                index = 0
            End If

            PictureBox1.Image.Save("Capture_B1.png")
        End If
        bmpBackground.Dispose()

#Region "Blackhole"

        'CaptureScreen()
        'PictureBox1.Image = bmpBackground
        'turn = SearchBitmap(bmpBackground, img(index), pt(index).X, pt(index).Y)
        'If (turn <> pt(index)) Then
        '    Me.Text = "AutoClick Searching ... Index : " + index.ToString() + " (False)"
        'Else
        '    Me.Text = "AutoClick Searching ... Index : " + index.ToString() + " (True) , SearchPoint = (" + turn.X.ToString() + " , " + turn.Y.ToString() + ")"

        '    If (index = 0) Then
        '        SendClick(index)
        '        Sleep(200)
        '        SendClick(index)
        '        Sleep(400)
        '        SendClick(index)
        '        Sleep(800)
        '        SendClickTracker_Track()
        '        index = 1
        '    ElseIf index = 1 Then
        '        SendClick(2) ' 2 = Press C
        '        index = 2
        '    ElseIf index = 2 Then
        '        SendClick(3)
        '        index = 0
        '    End If

        '    PictureBox1.Image.Save("Capture_B1.png")
        'End If
        'bmpBackground.Dispose()
#End Region

    End Sub

    Private Sub KeyboardTrigger_Tick(sender As System.Object, e As System.EventArgs) Handles KeyboardTrigger_Timer.Tick
        If GetAsyncKeyState(Keys.F1) Then
            Assistance_Timer.Enabled = True

        ElseIf GetAsyncKeyState(Keys.F2) Then
            Assistance_Timer.Enabled = False
            Me.Text = "AutoClick Stop."

        ElseIf GetAsyncKeyState(Keys.F3) Then 'Search Bitmap
            SearchIndexImage.Enabled = True

        ElseIf GetAsyncKeyState(Keys.F4) Then 'Capture winform Window
            CaptureImage.Enabled = True

        ElseIf GetAsyncKeyState(Keys.F5) Then
            index += 1
            Me.Text = "index = " + index.ToString()
        ElseIf GetAsyncKeyState(Keys.F6) Then
            index = 0
            Me.Text = "index = " + index.ToString()
        ElseIf GetAsyncKeyState(Keys.F7) Then
            'Simulate Press Key
            SendClick(index)

        End If
    End Sub

    Private Sub SearchIndexImage_Tick(sender As Object, e As EventArgs) Handles SearchIndexImage.Tick

        Dim SearchResult As Point

        CaptureScreen()
        PictureBox1.Image = bmpBackground

        SearchResult = SearchBitmap(bmpBackground, img(index), pt(index).X, pt(index).Y)
        Me.Text = "SearchResult = (" + SearchResult.X.ToString() + " , " + SearchResult.Y.ToString() + ")"

        bmpBackground.Dispose()

        SearchIndexImage.Stop()
    End Sub

    Private Sub CaptureImage_Tick(sender As Object, e As EventArgs) Handles CaptureImage.Tick

        index += 1
        Me.Text = index.ToString

        Dim str As String = DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss")
        Console.WriteLine("dt : " + str)
        CaptureScreen()
        bmpBackground.Save("CaptureScreen_" + str + ".png")

        Me.Text = Me.Text + " , " + str

        Sleep(100)

        bmpBackground = Nothing

        CaptureImage.Enabled = False
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick

        If TimerIndexCheck = index Then
            Console.WriteLine("TimerIndexCheck = " + TimerIndexCheck.ToString())
            SendClickTracker_Track()
        End If

        TimerIndexCheck = index
    End Sub

    Public Sub SendClickTracker_Track()

        If (IndexCounter > 0) Then

            Dim str As String = ""
            For i As Integer = 0 To IndexCounter - 1
                str += SendClickTracker(i).ToString() + "  "
            Next
            Console.WriteLine($"Offset Click , List ( Length = " + IndexCounter.ToString() + ") : " + str)

            Dim SendValue As Integer = 0
            Dim LastClickIndex As Integer = SendClickTracker(IndexCounter - 1)

            If TimerIndexCheck = 1 Then
                SendValue = 1
            ElseIf TimerIndexCheck = 2 Or TimerIndexCheck = 3 Then
                SendValue = 2
            ElseIf TimerIndexCheck = 4 Then
                SendValue = 3
            End If
            SendClick(SendValue)
            Sleep(500)

            If LastClickIndex = 1 Then
                SendValue = 1
            ElseIf LastClickIndex = 2 Or LastClickIndex = 3 Then
                SendValue = 2
            ElseIf LastClickIndex = 4 Then
                SendValue = 3
            End If
            SendClick(SendValue)

            'Console.WriteLine($"LastClickIndex : " + SendClickTracker(IndexCounter - 1).ToString() + " , SendValue Should be : " + SendValue.ToString())
            'Console.WriteLine()
            Sleep(100)

            IndexCounter = 0
        End If
    End Sub
End Class
