# TwinBasic 线程池类库

## Document

[https://doc.twinbasic.vb6.pro/en/tbman](https://doc.twinbasic.vb6.pro/en/tbman)

## 简介

这是一个为 TwinBasic 开发的高性能线程池类库，提供了简单易用但功能强大的多线程编程支持。通过线程池，您可以轻松管理和重用线程资源，避免频繁创建和销毁线程带来的开销。

## 主要特性

- **高效的线程管理**
  - 自动管理线程生命周期
  - 支持固定大小和自动扩展的线程池
  - 智能的任务队列管理

- **丰富的任务控制**
  - 支持任务优先级
  - 任务超时控制
  - 任务取消机制
  - 任务重试策略
  - 任务依赖关系

- **完善的错误处理**
  - 详细的错误状态追踪
  - 支持错误回调
  - 完整的错误日志记录

- **强大的扩展功能**
  - 任务完成回调
  - 性能监控统计
  - 自动负载均衡
  - 日志记录系统

## 核心组件

- `cThreadPool`: 线程池主类，负责管理线程和任务队列
- `cThread`: 任务类，封装了单个异步任务的所有操作和状态
- `mTask`: 工具模块，提供通用函数和辅助方法

## 目录

### 基础文档
1. [类库参考](./api-reference.md)
   - 详细的类、方法、属性文档
   - 参数说明和返回值
   - 使用注意事项

2. [使用教程](./tutorials.md)
   - 基础示例
   - 常见场景
   - 最佳实践

3. [示例代码](./examples.md)
   - 简单任务处理
   - HTTP下载示例
   - 批量任务处理
   - 任务依赖示例

### 高级功能
4. [高级特性](./advanced-features.md)
   - 任务优先级管理
   - 错误处理机制
   - 性能优化策略
   - 自动扩展功能

5. [任务数据与结果管理](./TaskDataAndResult.md)
   - TaskData 数据传递机制
   - TaskResult 结果存储
   - 类型安全的参数管理
   - 线程安全的数据访问

6. [性能优化指南](./PerformanceOptimization.md)
   - 数据传递优化
   - 内存管理策略
   - 线程池配置优化
   - 性能监控与调试

7. [特性集成指南](./FeatureIntegration.md)
   - 内置特性组合使用
   - 超时与重试机制集成
   - 任务取消与暂停功能
   - 日志记录系统

8. [高级应用示例](./AdvancedExamples.md)
   - 批量文件处理
   - 复杂业务场景
   - 性能优化实例
   - 错误处理最佳实践

### 参考资料
9. [故障排除指南](./Troubleshooting.md)
   - 常见问题诊断
   - 性能问题分析
   - 错误代码参考
   - 调试技巧

## 快速开始

### 最简示例
```vb
' 一句话创建并执行任务
New cThreadPool().Create(4).AddTask(AddressOf MyProc).WaitForCompletion
```

### 完整示例
```vb
' 创建线程池
Dim pool As New cThreadPool
pool.Create 4  ' 创建4个线程的线程池

' 添加任务
Dim task As cThread
Set task = pool.AddTask(AddressOf MyProc)

' 设置任务属性
task.SetTimeout 5000        ' 设置5秒超时
task.SetRetryPolicy 3, 1000 ' 最多重试3次，间隔1秒
task.SetOnComplete AddressOf TaskComplete
task.SetOnError AddressOf TaskError

' 等待任务完成
task.WaitForCompletion
```

## 系统要求

- TwinBasic 开发环境
- Windows操作系统
- 支持多线程的CPU

## 注意事项

1. 线程安全
   - 所有公共方法都是线程安全的
   - 回调函数在线程上下文中执行
   - 需要注意UI更新的线程同步

2. 资源管理
   - 合理设置线程池大小
   - 注意任务超时设置
   - 正确处理任务取消

3. 错误处理
   - 建议使用错误回调
   - 检查任务执行状态
   - 合理设置重试策略

## 许可

此类库根据许可证发布。详见 [LICENSE](./LICENSE) 文件。
