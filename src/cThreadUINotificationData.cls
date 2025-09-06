Class cThreadUINotificationData

    '@Author: 邓伟(woeoio)
    '@Email: 215879458@qq.com
    '@Description "线程UI通知数据封装类"

    ' 公共属性
    Public TaskId As String
    Public NotificationCode As Long
    Public DataType As Long  ' 0=Long, 1=String, 2=Double, 3=Boolean
    Public LongValue As Long
    Public StringValue As String
    Public DoubleValue As Double
    Public BooleanValue As Boolean

End Class
