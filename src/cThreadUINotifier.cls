Class cThreadUINotifier

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Description "UI线程通知管理器 - 负责处理子线程到UI线程的消息传递"

    ' Windows API 声明
    Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (ByRef pcWndClassEx As WNDCLASSEX) As Integer
    Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, _
        ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
        ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" ( _
        ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long

    ' 常量定义
    Private Const WM_USER As Long = &H400
    Private Const WM_THREAD_NOTIFY As Long = WM_USER + 1000
    Private Const WS_POPUP As Long = &H80000000
    Private Const WS_DISABLED As Long = &H8000000
    
    ' 窗口类结构
    Private Type WNDCLASSEX
        cbSize As Long
        style As Long
        lpfnWndProc As LongPtr
        cbClsExtra As Long
        cbWndExtra As Long
        hInstance As LongPtr
        hIcon As LongPtr
        hCursor As LongPtr
        hbrBackground As LongPtr
        lpszMenuName As LongPtr
        lpszClassName As LongPtr
        hIconSm As LongPtr
    End Type
    
    ' 通知数据结构（保留用于参考，实际使用 cThreadUINotificationData 类）
    Private Type NotificationData
        TaskId As String
        NotificationCode As Long
        DataType As Long  ' 0=Long, 1=String, 2=Double, 3=Boolean
        LongValue As Long
        StringValue As String
        DoubleValue As Double
        BooleanValue As Boolean
    End Type
    
    ' 成员变量
    Private m_MessageWindow As LongPtr
    Private m_ClassName As String
    Private m_hInstance As LongPtr
    Private m_IsInitialized As Boolean
    Private m_NotificationData As Collection  ' 存储通知数据的集合
    Private m_NotificationCounter As Long     ' 通知数据计数器
    
    ' 事件定义
    Public Event ThreadNotification(ByVal TaskId As String, ByVal NotificationCode As Long, ByVal Data As Variant)
    Public Event ThreadProgress(ByVal TaskId As String, ByVal CurrentValue As Long, ByVal MaxValue As Long, ByVal Message As String)
    Public Event ThreadCompleted(ByVal TaskId As String, ByVal Result As Variant)
    Public Event ThreadError(ByVal TaskId As String, ByVal ErrorCode As Long, ByVal ErrorMessage As String)
    
    ' 预定义通知代码
    Public Enum eNotificationCode
        Notify_Progress = 1
        Notify_Status = 2
        Notify_Warning = 3
        Notify_Error = 4
        Notify_Completed = 5
        Notify_Custom = 100  ' 用户自定义通知从100开始
    End Enum

    ' ======================
    ' 初始化与清理
    ' ======================
    
    Private Sub Class_Initialize()
        Set m_NotificationData = New Collection
        m_ClassName = "TBThreadNotifier_" & ObjPtr(Me)
        m_hInstance = GetModuleHandle(vbNullString)
        m_NotificationCounter = 0
        
        ' 自动初始化UI通知器
        AutoInitialize
    End Sub
    
    Private Sub Class_Terminate()
        Cleanup
    End Sub
    
    '@Description "初始化UI通知器（保留用于向后兼容）"
    Public Function Initialize() As Boolean
        ' 现在初始化在 Class_Initialize 中自动完成
        Initialize = m_IsInitialized
    End Function
    
    '@Description "自动初始化UI通知器"
    Private Sub AutoInitialize()
        If m_IsInitialized Then Exit Sub
        
        ' 注册窗口类
        Dim wc As WNDCLASSEX
        wc.cbSize = LenB(wc)
        wc.lpfnWndProc = GetAddressOf_MessageWindowProc()
        wc.hInstance = m_hInstance
        wc.lpszClassName = StrPtr(m_ClassName)
        
        If RegisterClassEx(wc) = 0 Then
            Exit Sub
        End If
        
        ' 创建隐藏的消息窗口
        m_MessageWindow = CreateWindowEx(0, m_ClassName, "ThreadNotifier", _
                                       WS_POPUP Or WS_DISABLED, 0, 0, 0, 0, _
                                       0, 0, m_hInstance, 0)
        
        If m_MessageWindow = 0 Then
            UnregisterClass m_ClassName, m_hInstance
            Exit Sub
        End If
        
        m_IsInitialized = True
    End Sub
    
    '@Description "清理资源"
    Private Sub Cleanup()
        If m_MessageWindow <> 0 Then
            DestroyWindow m_MessageWindow
            m_MessageWindow = 0
        End If
        
        If m_ClassName <> "" Then
            UnregisterClass m_ClassName, m_hInstance
            m_ClassName = ""
        End If
        
        Set m_NotificationData = Nothing
        m_IsInitialized = False
    End Sub

    ' ======================
    ' 公共接口
    ' ======================
    
    '@Description "获取消息窗口句柄，供线程使用"
    Public Property Get MessageWindowHandle() As LongPtr
        MessageWindowHandle = m_MessageWindow
    End Property
    
    '@Description "检查UI通知器是否已成功初始化"
    Public Property Get IsInitialized() As Boolean
        IsInitialized = m_IsInitialized
    End Property
    
    '@Description "发送通知到UI线程（供子线程调用）"
    Public Sub PostNotification(ByVal TaskId As String, ByVal NotificationCode As Long, Optional ByVal Data As Variant)
        If Not m_IsInitialized Then Exit Sub
        
        ' 创建通知数据对象
        Dim notifyData As New cThreadUINotificationData
        notifyData.TaskId = TaskId
        notifyData.NotificationCode = NotificationCode
        
        ' 根据数据类型存储
        If Not IsMissing(Data) Then
            Select Case VarType(Data)
                Case vbLong, vbInteger, vbByte
                    notifyData.DataType = 0
                    notifyData.LongValue = CLng(Data)
                Case vbString
                    notifyData.DataType = 1
                    notifyData.StringValue = CStr(Data)
                Case vbDouble, vbSingle, vbCurrency
                    notifyData.DataType = 2
                    notifyData.DoubleValue = CDbl(Data)
                Case vbBoolean
                    notifyData.DataType = 3
                    notifyData.BooleanValue = CBool(Data)
                Case Else
                    notifyData.DataType = 1
                    notifyData.StringValue = CStr(Data)
            End Select
        End If
        
        ' 存储到集合中（使用递增的数字ID作为键）
        m_NotificationCounter = m_NotificationCounter + 1
        Dim dataKey As String
        dataKey = "NotifyData_" & m_NotificationCounter
        m_NotificationData.Add notifyData, dataKey
        
        ' 发送消息到UI线程（使用数字ID）
        PostMessage m_MessageWindow, WM_THREAD_NOTIFY, m_NotificationCounter, ObjPtr(Me)
    End Sub
    
    '@Description "发送进度通知"
    Public Sub PostProgress(ByVal TaskId As String, ByVal CurrentValue As Long, ByVal MaxValue As Long, Optional ByVal Message As String = "")
        Dim progressInfo As String
        progressInfo = CurrentValue & "|" & MaxValue & "|" & Message
        PostNotification TaskId, Notify_Progress, progressInfo
    End Sub
    
    '@Description "发送完成通知"
    Public Sub PostCompleted(ByVal TaskId As String, Optional ByVal Result As Variant)
        PostNotification TaskId, Notify_Completed, Result
    End Sub
    
    '@Description "发送错误通知"
    Public Sub PostError(ByVal TaskId As String, ByVal ErrorCode As Long, ByVal ErrorMessage As String)
        Dim errorInfo As String
        errorInfo = ErrorCode & "|" & ErrorMessage
        PostNotification TaskId, Notify_Error, errorInfo
    End Sub

    ' ======================
    ' 内部方法
    ' ======================
    
    '@Description "获取消息处理过程的地址"
    Private Function GetAddressOf_MessageWindowProc() As LongPtr
        ' 在 TwinBasic 中使用 AddressOf 操作符
        GetAddressOf_MessageWindowProc = AddressOf MessageWindowProc
    End Function
    
    '@Description "消息窗口过程"
    Private Function MessageWindowProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                                     ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
        If uMsg = WM_THREAD_NOTIFY Then
            ' 处理线程通知
            HandleThreadNotification wParam, lParam
            MessageWindowProc = 0
        Else
            MessageWindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
        End If
    End Function
    
    '@Description "处理线程通知"
    Private Sub HandleThreadNotification(ByVal wParam As LongPtr, ByVal lParam As LongPtr)
        On Error GoTo ErrorHandler
        
        ' 从 wParam 获取数据ID
        Dim dataId As Long
        dataId = CLng(wParam)
        
        ' 构造数据键
        Dim dataKey As String
        dataKey = "NotifyData_" & dataId
        
        ' 从集合中获取通知数据对象
        Dim notifyData As cThreadUINotificationData
        Set notifyData = m_NotificationData(dataKey)
        
        ' 如果没有找到数据，退出
        If notifyData Is Nothing Then Exit Sub
        
        ' 移除已处理的数据
        m_NotificationData.Remove dataKey
        
        ' 转换数据
        Dim data As Variant
        Select Case notifyData.DataType
            Case 0: data = notifyData.LongValue
            Case 1: data = notifyData.StringValue
            Case 2: data = notifyData.DoubleValue
            Case 3: data = notifyData.BooleanValue
            Case Else: data = Null
        End Select
        
        ' 根据通知类型触发相应事件
        Select Case notifyData.NotificationCode
            Case Notify_Progress
                ' 解析进度信息
                Dim parts() As String
                parts = Split(CStr(data), "|")
                If UBound(parts) >= 1 Then
                    RaiseEvent ThreadProgress(notifyData.TaskId, CLng(parts(0)), CLng(parts(1)), _
                                            IIf(UBound(parts) >= 2, parts(2), ""))
                End If
                
            Case Notify_Completed
                RaiseEvent ThreadCompleted(notifyData.TaskId, data)
                
            Case Notify_Error
                ' 解析错误信息
                Dim errorParts() As String
                errorParts = Split(CStr(data), "|")
                If UBound(errorParts) >= 1 Then
                    RaiseEvent ThreadError(notifyData.TaskId, CLng(errorParts(0)), errorParts(1))
                End If
                
            Case Else
                ' 通用通知
                RaiseEvent ThreadNotification(notifyData.TaskId, notifyData.NotificationCode, data)
        End Select
        
        Exit Sub
        
    ErrorHandler:
        ' 如果发生错误（比如找不到指定的键），静默处理
        ' 这可能发生在消息处理延迟的情况下
        Exit Sub
    End Sub

End Class
