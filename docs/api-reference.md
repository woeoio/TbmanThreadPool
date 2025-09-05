# API 参考文档

## cTasks 类

线程池主类，负责管理线程和任务队列。

### 属性

#### ThreadCount
- **类型**: Long
- **说明**: 获取当前线程池中的线程数量
- **访问级别**: 只读
- **示例**:
```vb
Debug.Print "当前线程数：" & pool.ThreadCount
```

#### ActiveThreadCount
- **类型**: Long
- **说明**: 获取当前正在执行任务的线程数量
- **访问级别**: 只读
- **示例**:
```vb
Debug.Print "活动线程数：" & pool.ActiveThreadCount
```

#### PendingTaskCount
- **类型**: Long
- **说明**: 获取等待执行的任务数量
- **访问级别**: 只读
- **示例**:
```vb
Debug.Print "等待任务数：" & pool.PendingTaskCount
```

### 方法

#### Create
```vb
Public Function Create(ByVal threadCount As Long) As Boolean
```
- **说明**: 创建指定数量的线程池
- **参数**: 
  - threadCount: 要创建的线程数量
- **返回值**: 创建成功返回 True
- **示例**:
```vb
Dim pool As New cTasks
pool.Create 4  ' 创建4个线程的池
```

#### AddTask
```vb
Public Function AddTask(ByVal proc As LongPtr, _
                       Optional ByVal param As Variant, _
                       Optional ByVal priority As TaskPriority = Priority_Normal) As cTask
```
- **说明**: 添加新任务到线程池
- **参数**:
  - proc: 任务过程的地址
  - param: 传递给任务的参数(可选)
  - priority: 任务优先级(可选)
- **返回值**: 新创建的任务对象
- **示例**:
```vb
Dim task As cTask
Set task = pool.AddTask(AddressOf MyProc, "参数", Priority_High)
```

#### WaitForAll
```vb
Public Function WaitForAll(Optional ByVal timeout As Long = INFINITE) As Boolean
```
- **说明**: 等待所有任务完成
- **参数**:
  - timeout: 超时时间(毫秒)，默认无限等待
- **返回值**: 所有任务完成返回 True
- **示例**:
```vb
' 等待所有任务完成，最多等待10秒
If Not pool.WaitForAll(10000) Then
    Debug.Print "有任务未完成"
End If
```

#### Shutdown
```vb
Public Sub Shutdown(Optional ByVal graceful As Boolean = True)
```
- **说明**: 关闭线程池
- **参数**:
  - graceful: 是否优雅关闭，True则等待任务完成
- **示例**:
```vb
' 优雅关闭，等待任务完成
pool.Shutdown True

' 强制关闭，不等待任务
pool.Shutdown False
```

#### SetAutoScale
```vb
Public Sub SetAutoScale(ByVal enabled As Boolean, _
                       Optional ByVal minThreads As Long = 2, _
                       Optional ByVal loadFactor As Double = 0.75)
```
- **说明**: 设置线程池自动扩展
- **参数**:
  - enabled: 是否启用自动扩展
  - minThreads: 最小线程数
  - loadFactor: 负载因子(0-1)
- **示例**:
```vb
' 启用自动扩展，最小2线程，负载因子0.75
pool.SetAutoScale True, 2, 0.75
```

## cTask 类

单个任务的封装类，提供任务控制和状态管理。

### 属性

#### Status
- **类型**: TaskStatus 枚举
- **说明**: 获取任务当前状态
- **可能值**:
  - TaskStatus_Created: 已创建
  - TaskStatus_Running: 正在运行
  - TaskStatus_Completed: 已完成
  - TaskStatus_Failed: 失败
  - TaskStatus_Cancelled: 已取消
  - TaskStatus_TimedOut: 已超时

#### IsRunning
- **类型**: Boolean
- **说明**: 任务是否正在运行

#### IsCompleted
- **类型**: Boolean
- **说明**: 任务是否已完成

#### Result
- **类型**: Long
- **说明**: 获取任务执行结果

### 方法

#### Create
```vb
Public Function Create(ByVal proc As LongPtr, _
                      Optional ByVal param As Variant) As Boolean
```
- **说明**: 创建新任务
- **参数**:
  - proc: 任务过程地址
  - param: 任务参数(可选)
- **返回值**: 创建成功返回 True

#### SetTimeout
```vb
Public Sub SetTimeout(ByVal milliseconds As Long)
```
- **说明**: 设置任务超时时间
- **参数**:
  - milliseconds: 超时时间(毫秒)

#### SetRetryPolicy
```vb
Public Sub SetRetryPolicy(ByVal maxRetries As Long, _
                         ByVal retryInterval As Long)
```
- **说明**: 设置重试策略
- **参数**:
  - maxRetries: 最大重试次数
  - retryInterval: 重试间隔(毫秒)

#### SetOnComplete
```vb
Public Sub SetOnComplete(ByVal callback As LongPtr)
```
- **说明**: 设置完成回调
- **参数**:
  - callback: 回调函数地址

#### SetOnError
```vb
Public Sub SetOnError(ByVal callback As LongPtr)
```
- **说明**: 设置错误回调
- **参数**:
  - callback: 回调函数地址

#### WaitForCompletion
```vb
Public Function WaitForCompletion(Optional ByVal timeout As Long = INFINITE) As Boolean
```
- **说明**: 等待任务完成
- **参数**:
  - timeout: 超时时间(毫秒)
- **返回值**: 任务完成返回 True

#### RequestCancel
```vb
Public Sub RequestCancel()
```
- **说明**: 请求取消任务
