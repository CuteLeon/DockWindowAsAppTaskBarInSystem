Public Class DockForm
    Public Declare Function SHAppBarMessage Lib "Shell32.dll" (ByVal dwMessage As Integer, ByRef pData As APPBARDATA) As Integer
    Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uiAction As Integer, ByVal uiParam As Integer, ByRef pvParam As RECT, ByVal fWinIni As Integer) As Integer

    Public Const HWND_TOPMOST = -1
    Public Const ABE_LEFT = 0
    Public Const ABE_TOP = 1
    Public Const ABE_RIGHT = 2
    Public Const ABE_BOTTOM = 3
    Public Const ABM_NEW = &H0 '注册一个系统任务栏
    Public Const ABM_REMOVE = &H1 '移除一个系统任务栏
    Public Const ABM_SETPOS = &H3 '设置任务栏的尺寸和位置
    Public Const ABM_GETSTATE = &H4 '获取任务栏是否自动隐藏和置顶显示
    Public Const ABM_GETTASKBARPOS = &H5 '获取任务栏位置信息
    Public Const ABM_ACTIVATE = &H6 '激活或失活一个任务栏，pData的lParam设置为True为激活，False为失活
    Public Const ABM_GETAUTOHIDEBAR = &H7 '检索一个与屏幕指定边界相关的自动隐藏句柄
    Public Const ABM_SETAUTOHIDEBAR = &H8 '注册一个与屏幕指定边界相关的自动隐藏任务栏
    Public Const ABM_WINDOWPOSCHANGED = &H9 '通知系统一个任务栏的位置发生了改变
    Public Const ABM_SETSTATE = &HA '设置一个任务栏的自动隐藏和置顶显示属性(XP及以后)
    Public Const ABM_GETAUTOHIDEBAREX = &HB '检索一个与屏幕指定边界相关的自动隐藏句柄(XP及以后)
    Public Const ABM_SETAUTOHIDEBAREX = &HC '注册一个与屏幕指定边界相关的自动隐藏任务栏(XP及以后)
    Public Const SPI_GETWORKAREA = 48
    Public Const WM_MOUSEMOVE = &H200
    Public Const SWP_SHOWWINDOW = &H40

    Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    Public Structure APPBARDATA
        Dim cbSize As Integer
        Dim hwnd As Integer
        Dim uCallbackMessage As Integer
        Dim uEdge As Integer
        Dim rc As RECT
        Dim lParam As Integer
    End Structure

    Public Enum DockSide
        dsLeft = 0
        dsTop = 1
        dsRight = 2
        dsBottom = 3
    End Enum

    Dim AppBar As APPBARDATA '代表停靠窗口

    Public Sub myDock(ByVal ScreenSide As DockSide, ByRef xForm As Form, ByRef AppBar As APPBARDATA)
        Dim ScreenSize As Size = My.Computer.Screen.Bounds.Size  ' 屏幕尺寸
        Dim DockSize As Size = xForm.Size '停靠区域尺寸
        Dim bResult As Boolean 'API返回结果
        Dim WorkArea As RECT  '工作区域

        AppBar.hwnd = xForm.Handle '停靠窗体的句柄
        AppBar.cbSize = Len(AppBar) ' 停靠窗口需要字节数
        AppBar.uCallbackMessage = WM_MOUSEMOVE '对任何系统消息调用返回功能

        bResult = SHAppBarMessage(ABM_REMOVE, AppBar) '如果已经停靠，先取消停靠
        My.Application.DoEvents()  '处理堆积的消息

        SystemParametersInfo(SPI_GETWORKAREA, 0, WorkArea, 0) '获取尚未停靠的屏幕区域
        Debug.Print(WorkArea.Left & "," & WorkArea.Top & "," & WorkArea.Right & "," & WorkArea.Bottom)

        bResult = SHAppBarMessage(ABM_NEW, AppBar) '注册停靠窗体

        Select Case ScreenSide '判断停靠方向
            Case ABE_TOP '顶部
                AppBar.uEdge = ABE_TOP
                AppBar.rc.Top = WorkArea.Top
                AppBar.rc.Left = 0
                AppBar.rc.Right = ScreenSize.Width
                AppBar.rc.Bottom = DockSize.Height
            Case ABE_BOTTOM '底部
                AppBar.uEdge = ABE_BOTTOM
                AppBar.rc.Top = WorkArea.Bottom - DockSize.Height
                AppBar.rc.Left = 0
                AppBar.rc.Right = ScreenSize.Width
                AppBar.rc.Bottom = WorkArea.Bottom
            Case ABE_LEFT '左部
                AppBar.uEdge = ABE_LEFT
                AppBar.rc.Top = 0
                AppBar.rc.Left = WorkArea.Left
                AppBar.rc.Right = WorkArea.Left + DockSize.Width
                AppBar.rc.Bottom = ScreenSize.Height
            Case ABE_RIGHT '右部
                AppBar.uEdge = ABE_RIGHT
                AppBar.rc.Top = 0
                AppBar.rc.Left = WorkArea.Right - DockSize.Width
                AppBar.rc.Right = WorkArea.Right
                AppBar.rc.Bottom = ScreenSize.Height
        End Select

        bResult = SHAppBarMessage(ABM_SETPOS, AppBar) '保留空闲区域给停靠窗口
        My.Application.DoEvents() '处理堆积消息

        '窗口置前，修改窗体位置和大小
        bResult = SetWindowPos(xForm.Handle, HWND_TOPMOST, AppBar.rc.Left, AppBar.rc.Top, AppBar.rc.Right, AppBar.rc.Bottom, SWP_SHOWWINDOW)
    End Sub

    Public Sub UnDock(ByRef AppBar As APPBARDATA)
        Call SHAppBarMessage(ABM_REMOVE, AppBar) '解除停靠
        My.Application.DoEvents()
    End Sub

    Public Sub ResetSize()
        '将停靠窗口恢复到屏幕中心
        Me.Height = 100
        Me.Width = 100
        Me.Left = (My.Computer.Screen.Bounds.Width - Me.Width) / 2
        Me.Top = (My.Computer.Screen.Bounds.Height - Me.Height) / 2
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myDock(DockSide.dsLeft, Me, AppBar)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        UnDock(AppBar)
        ResetSize()
    End Sub

    Private Sub DockForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '停靠任务栏的背景是桌面壁纸，所以把窗体背景
        ResetSize()
    End Sub
End Class
