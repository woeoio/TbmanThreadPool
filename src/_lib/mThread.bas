Module mThread

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Document: https://doc.twinbasic.vb6.pro/en/tbman/threadPool/

    ' Windows API 声明
    Public Declare PtrSafe Function CreateThread Lib "kernel32" ( _
        ByVal lpThreadAttributes As LongPtr, _
        ByVal dwStackSize As LongPtr, _
        ByVal lpStartAddress As LongPtr, _
        ByVal lpParameter As LongPtr, _
        ByVal dwCreationFlags As Long, _
        ByRef lpThreadId As Long _
    ) As LongPtr

    Public Declare PtrSafe Function SuspendThread Lib "kernel32" ( _
        ByVal hThread As LongPtr _
    ) As Long

    Public Declare PtrSafe Function ResumeThread Lib "kernel32" ( _
        ByVal hThread As LongPtr _
    ) As Long

    Public Declare PtrSafe Function TerminateThread Lib "kernel32" ( _
        ByVal hThread As LongPtr, _
        ByVal dwExitCode As Long _
    ) As Long

    Public Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As LongPtr, _
        ByVal dwMilliseconds As Long _
    ) As Long

    Public Declare PtrSafe Function WaitForMultipleObjects Lib "kernel32" ( _
        ByVal nCount As Long, _
        ByRef lpHandles As LongPtr, _
        ByVal bWaitAll As Long, _
        ByVal dwMilliseconds As Long _
    ) As Long

    Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr _
    ) As Long

    Public Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long _
    )

    ' 常量定义
    Public Const INFINITE As Long = &HFFFFFFFF
    Public Const WAIT_FAILED As Long = &HFFFFFFFF
    Public Const WAIT_OBJECT_0 As Long = 0
    Public Const WAIT_TIMEOUT As Long = 258

    ' 内存操作 API (用于 ReturnFromPtr)
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, _
        ByRef Source As Any, _
        ByVal Length As LongPtr _
    )
    
    ' ' 辅助函数：从指针获取对象
    Public Function ReturnFromPtr(ByVal ptr As LongPtr) As Object
        Dim obj As Object
        ' 注意：这需要 TwinBasic 支持或适当的 COM 接口处理
        CopyMemory obj, ptr, LenB(ptr)
        Set ReturnFromPtr = obj
        Set obj = Nothing
    End Function    
    ' Public Function ObjectFromPtr(ByVal ptr As LongPtr) As Object
    '     Dim obj As Object
    '     ' 注意：这需要 TwinBasic 支持或适当的 COM 接口处理
    '     CopyMemory obj, ptr, LenB(ptr)
    '     Set ObjectFromPtr = obj
    '     Set obj = Nothing
    ' End Function

    ' ' 辅助函数：获取对象指针，可直接使用 ObjPtr，无需此函数
    ' Public Function PtrFromObject(ByVal obj As Object) As LongPtr
    '     PtrFromObject = ObjPtr(obj)
    ' End Function



End Module