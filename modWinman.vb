Imports System.Runtime.InteropServices

Module modWinman

    ' Declare Windows API functions
    <DllImport("user32.dll", SetLastError:=True)>
    Function FindWindow(ByVal lpClassName As String,
                    ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Function ShowWindow(ByVal hWnd As IntPtr,
                    ByVal nCmdShow As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Function SetWindowLong(ByVal hWnd As IntPtr,
                       ByVal nIndex As Integer,
                       ByVal dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Function SetWindowPos(ByVal hWnd As IntPtr,
                      ByVal hWndInsertAfter As IntPtr,
                      ByVal X As Integer,
                      ByVal Y As Integer,
                      ByVal cx As Integer,
                      ByVal cy As Integer,
                      ByVal uFlags As UInteger) As Boolean
    End Function

    ' Define constants
    Const WM_SETREDRAW = &HB
    Const GWL_STYLE = -16
    Const WS_VISIBLE = &H10000000
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const SWP_NOZORDER = &H4
    Const SWP_FRAMECHANGED = &H20

    ' Define constants for SW_HIDE and SW_SHOW
    Const SW_HIDE As Integer = 0
    Const SW_SHOW = 5

    ' Define constants for WM_CLOSE
    Const WM_CLOSE As UInteger = &H10

    Sub HideHiddenWindow()
        ' Find the window handle by class name and window name
        Dim hWnd = FindWindow("MSNHiddenWindowClass", Nothing)
        If hWnd = IntPtr.Zero Then
            Return
        End If

        ' Method 1: Use WM_SETREDRAW message with wParam set to FALSE
        SendMessage(hWnd, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero)

    End Sub

    Sub ShowHiddenWindow()
        ' Find the window handle by class name and window name
        Dim hWnd = FindWindow("MSNHiddenWindowClass", Nothing)
        If hWnd = IntPtr.Zero Then
            Return
        End If

        ShowWindow(hWnd, SW_SHOW)
    End Sub

    Sub ShowWLMWindow()
        Process.Start(My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\Windows Live\Messenger\msnmsgr.exe")
    End Sub

    Sub CloseWLMWindow()
        Process.Start(My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\Windows Live\Messenger\msnmsgr.exe")
    End Sub

    Sub CloseHiddenWindow()
        ' Find the window handle by class name and window name
        Dim hWnd = FindWindow("MSNHiddenWindowClass", Nothing)
        If hWnd = IntPtr.Zero Then
            Return
        End If

        SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)
    End Sub

    ' based upon Craftplacer's work https://github.com/Craftplacer/MessengerDotNet/blob/master/MessengerActivity.cs
    Private Function VarPtr(ByVal e As Object) As IntPtr
        Dim handle = GCHandle.Alloc(e, GCHandleType.Pinned)
        Dim ptr As IntPtr = handle.AddrOfPinnedObject()
        handle.Free()
        Return ptr
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure COPYDATASTRUCT
        Public dwData As IntPtr
        Public cbData As Integer
        Public lpData As IntPtr
    End Structure

    <DllImport("user32.dll", EntryPoint:="FindWindowExA")>
    Private Function FindWindowEx(ByVal hWnd1 As IntPtr, ByVal hWnd2 As IntPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As IntPtr
    End Function

    Sub setActivity(kind As String, text As String)
        Dim format As String = String.Format("\\0{0}\\01\\0{1}\\0\0", kind, text)
        Dim data = New COPYDATASTRUCT With {
            .dwData = CType(&H547, IntPtr),
            .lpData = VarPtr(format),
            .cbData = format.Length * 2
        }
        Dim ptr As IntPtr = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "MsnMsgrUIManager", Nothing)

        If ptr.ToInt32() > 0 Then
            SendMessage(ptr, &H4A, IntPtr.Zero, VarPtr(data))
        End If
    End Sub

End Module
