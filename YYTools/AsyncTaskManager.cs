using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YYTools
{
    /// <summary>
    /// 异步任务管理器
    /// </summary>
    public class AsyncTaskManager : IDisposable
    {
        private readonly Dictionary<string, CancellationTokenSource> _taskTokens = new Dictionary<string, CancellationTokenSource>();
        private readonly Dictionary<string, Task> _runningTasks = new Dictionary<string, Task>();
        private bool _disposed = false;

        #region 事件

        /// <summary>
        /// 任务进度报告事件
        /// </summary>
        public event EventHandler<TaskProgressEventArgs> ProgressReported;

        /// <summary>
        /// 任务完成事件
        /// </summary>
        public event EventHandler<TaskCompletedEventArgs> TaskCompleted;

        /// <summary>
        /// 任务取消事件
        /// </summary>
        public event EventHandler<TaskCancelledEventArgs> TaskCancelled;

        /// <summary>
        /// 任务错误事件
        /// </summary>
        public event EventHandler<TaskErrorEventArgs> TaskError;

        #endregion

        #region 任务管理

        /// <summary>
        /// 启动异步任务
        /// </summary>
        public async Task<TaskResult<T>> StartTaskAsync<T>(
            string taskName,
            Func<CancellationToken, IProgress<TaskProgress>, Task<T>> taskFactory,
            bool allowMultiple = false)
        {
            try
            {
                // 检查是否允许重复任务
                if (!allowMultiple && _runningTasks.ContainsKey(taskName))
                {
                    throw new InvalidOperationException($"任务 '{taskName}' 已在运行中");
                }

                // 创建取消令牌
                var cancellationTokenSource = new CancellationTokenSource();
                var progress = new TaskProgressReporter(this, taskName);

                // 记录任务开始
                Logger.LogUserAction("启动异步任务", $"任务名称: {taskName}", "开始");
                Logger.LogPerformance($"任务启动: {taskName}", TimeSpan.Zero);

                // 创建并启动任务
                var task = Task.Run(async () =>
                {
                    try
                    {
                        var result = await taskFactory(cancellationTokenSource.Token, progress);
                        return new TaskResult<T> { Success = true, Data = result };
                    }
                    catch (OperationCanceledException)
                    {
                        Logger.LogUserAction("任务被取消", $"任务名称: {taskName}", "已取消");
                        OnTaskCancelled(taskName, "任务被用户取消");
                        return new TaskResult<T> { Success = false, IsCancelled = true };
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError($"任务执行失败: {taskName}", ex);
                        OnTaskError(taskName, ex);
                        return new TaskResult<T> { Success = false, Error = ex };
                    }
                }, cancellationTokenSource.Token);

                // 注册任务
                _taskTokens[taskName] = cancellationTokenSource;
                _runningTasks[taskName] = task;

                // 等待任务完成
                var taskResult = await task;

                // 清理任务
                CleanupTask(taskName);

                // 记录任务完成
                Logger.LogUserAction("异步任务完成", $"任务名称: {taskName}", taskResult.Success ? "成功" : "失败");
                OnTaskCompleted(taskName, taskResult);

                return taskResult;
            }
            catch (Exception ex)
            {
                Logger.LogError($"启动异步任务失败: {taskName}", ex);
                throw;
            }
        }

        /// <summary>
        /// 启动后台任务（不等待完成）
        /// </summary>
        public void StartBackgroundTask(
            string taskName,
            Func<CancellationToken, IProgress<TaskProgress>, Task> taskFactory,
            bool allowMultiple = false)
        {
            try
            {
                // 检查是否允许重复任务
                if (!allowMultiple && _runningTasks.ContainsKey(taskName))
                {
                    Logger.LogWarning($"任务 '{taskName}' 已在运行中，跳过启动");
                    return;
                }

                // 创建取消令牌
                var cancellationTokenSource = new CancellationTokenSource();
                var progress = new TaskProgressReporter(this, taskName);

                // 记录任务开始
                Logger.LogUserAction("启动后台任务", $"任务名称: {taskName}", "开始");

                // 创建并启动任务
                var task = Task.Run(async () =>
                {
                    try
                    {
                        await taskFactory(cancellationTokenSource.Token, progress);
                        Logger.LogUserAction("后台任务完成", $"任务名称: {taskName}", "成功");
                    }
                    catch (OperationCanceledException)
                    {
                        Logger.LogUserAction("后台任务被取消", $"任务名称: {taskName}", "已取消");
                        OnTaskCancelled(taskName, "任务被用户取消");
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError($"后台任务执行失败: {taskName}", ex);
                        OnTaskError(taskName, ex);
                    }
                    finally
                    {
                        CleanupTask(taskName);
                    }
                }, cancellationTokenSource.Token);

                // 注册任务
                _taskTokens[taskName] = cancellationTokenSource;
                _runningTasks[taskName] = task;
            }
            catch (Exception ex)
            {
                Logger.LogError($"启动后台任务失败: {taskName}", ex);
                throw;
            }
        }

        /// <summary>
        /// 取消任务
        /// </summary>
        public bool CancelTask(string taskName)
        {
            try
            {
                if (_taskTokens.TryGetValue(taskName, out var tokenSource))
                {
                    tokenSource.Cancel();
                    Logger.LogUserAction("取消任务", $"任务名称: {taskName}", "已发送取消信号");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Logger.LogError($"取消任务失败: {taskName}", ex);
                return false;
            }
        }

        /// <summary>
        /// 取消所有任务
        /// </summary>
        public void CancelAllTasks()
        {
            try
            {
                var taskNames = new List<string>(_taskTokens.Keys);
                foreach (var taskName in taskNames)
                {
                    CancelTask(taskName);
                }
                Logger.LogUserAction("取消所有任务", "批量操作", $"已发送取消信号给 {taskNames.Count} 个任务");
            }
            catch (Exception ex)
            {
                Logger.LogError("取消所有任务失败", ex);
            }
        }

        /// <summary>
        /// 检查任务是否正在运行
        /// </summary>
        public bool IsTaskRunning(string taskName)
        {
            return _runningTasks.ContainsKey(taskName) && !_runningTasks[taskName].IsCompleted;
        }

        /// <summary>
        /// 获取运行中的任务列表
        /// </summary>
        public List<string> GetRunningTasks()
        {
            var runningTasks = new List<string>();
            foreach (var kvp in _runningTasks)
            {
                if (!kvp.Value.IsCompleted)
                {
                    runningTasks.Add(kvp.Key);
                }
            }
            return runningTasks;
        }

        /// <summary>
        /// 等待任务完成
        /// </summary>
        public async Task WaitForTaskAsync(string taskName, TimeSpan timeout = default)
        {
            try
            {
                if (_runningTasks.TryGetValue(taskName, out var task))
                {
                    if (timeout == default)
                    {
                        await task;
                    }
                    else
                    {
                        using (var timeoutCts = new CancellationTokenSource(timeout))
                        {
                            await Task.WhenAny(task, Task.Delay(-1, timeoutCts.Token));
                            if (!task.IsCompleted)
                            {
                                throw new TimeoutException($"任务 '{taskName}' 超时");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"等待任务完成失败: {taskName}", ex);
                throw;
            }
        }

        #endregion

        #region 进度报告

        /// <summary>
        /// 报告任务进度
        /// </summary>
        internal void ReportProgress(string taskName, TaskProgress progress)
        {
            try
            {
                OnProgressReported(taskName, progress);
            }
            catch (Exception ex)
            {
                Logger.LogError($"报告任务进度失败: {taskName}", ex);
            }
        }

        #endregion

        #region 事件触发

        protected virtual void OnProgressReported(string taskName, TaskProgress progress)
        {
            ProgressReported?.Invoke(this, new TaskProgressEventArgs(taskName, progress));
        }

        protected virtual void OnTaskCompleted(string taskName, object result)
        {
            TaskCompleted?.Invoke(this, new TaskCompletedEventArgs(taskName, result));
        }

        protected virtual void OnTaskCancelled(string taskName, string reason)
        {
            TaskCancelled?.Invoke(this, new TaskCancelledEventArgs(taskName, reason));
        }

        protected virtual void OnTaskError(string taskName, Exception error)
        {
            TaskError?.Invoke(this, new TaskErrorEventArgs(taskName, error));
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 清理任务资源
        /// </summary>
        private void CleanupTask(string taskName)
        {
            try
            {
                if (_taskTokens.TryGetValue(taskName, out var tokenSource))
                {
                    _taskTokens.Remove(taskName);
                    tokenSource.Dispose();
                }

                if (_runningTasks.TryGetValue(taskName, out var task))
                {
                    _runningTasks.Remove(taskName);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"清理任务资源失败: {taskName}", ex);
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!_disposed)
            {
                try
                {
                    CancelAllTasks();
                    
                    // 等待所有任务完成
                    var tasks = new List<Task>(_runningTasks.Values);
                    if (tasks.Count > 0)
                    {
                        Task.WaitAll(tasks.ToArray(), TimeSpan.FromSeconds(10));
                    }

                    // 清理资源
                    foreach (var tokenSource in _taskTokens.Values)
                    {
                        tokenSource.Dispose();
                    }
                    _taskTokens.Clear();
                    _runningTasks.Clear();
                }
                catch (Exception ex)
                {
                    Logger.LogError("释放异步任务管理器资源失败", ex);
                }
                finally
                {
                    _disposed = true;
                }
            }
        }

        #endregion
    }

    #region 任务进度报告器

    /// <summary>
    /// 任务进度报告器
    /// </summary>
    internal class TaskProgressReporter : IProgress<TaskProgress>
    {
        private readonly AsyncTaskManager _taskManager;
        private readonly string _taskName;

        public TaskProgressReporter(AsyncTaskManager taskManager, string taskName)
        {
            _taskManager = taskManager;
            _taskName = taskName;
        }

        public void Report(TaskProgress value)
        {
            _taskManager.ReportProgress(_taskName, value);
        }
    }

    #endregion

    #region 事件参数类

    /// <summary>
    /// 任务进度事件参数
    /// </summary>
    public class TaskProgressEventArgs : EventArgs
    {
        public string TaskName { get; }
        public TaskProgress Progress { get; }

        public TaskProgressEventArgs(string taskName, TaskProgress progress)
        {
            TaskName = taskName;
            Progress = progress;
        }
    }

    /// <summary>
    /// 任务完成事件参数
    /// </summary>
    public class TaskCompletedEventArgs : EventArgs
    {
        public string TaskName { get; }
        public object Result { get; }

        public TaskCompletedEventArgs(string taskName, object result)
        {
            TaskName = taskName;
            Result = result;
        }
    }

    /// <summary>
    /// 任务取消事件参数
    /// </summary>
    public class TaskCancelledEventArgs : EventArgs
    {
        public string TaskName { get; }
        public string Reason { get; }

        public TaskCancelledEventArgs(string taskName, string reason)
        {
            TaskName = taskName;
            Reason = reason;
        }
    }

    /// <summary>
    /// 任务错误事件参数
    /// </summary>
    public class TaskErrorEventArgs : EventArgs
    {
        public string TaskName { get; }
        public Exception Error { get; }

        public TaskErrorEventArgs(string taskName, Exception error)
        {
            TaskName = taskName;
            Error = error;
        }
    }

    #endregion

    #region 任务结果类

    /// <summary>
    /// 任务结果
    /// </summary>
    public class TaskResult<T>
    {
        public bool Success { get; set; }
        public T Data { get; set; }
        public bool IsCancelled { get; set; }
        public Exception Error { get; set; }
        public string Message { get; set; }
    }

    #endregion

    #region 任务进度类

    /// <summary>
    /// 任务进度
    /// </summary>
    public class TaskProgress
    {
        public int Percentage { get; set; }
        public string Message { get; set; }
        public string CurrentOperation { get; set; }
        public int CurrentStep { get; set; }
        public int TotalSteps { get; set; }
        public TimeSpan ElapsedTime { get; set; }
        public TimeSpan EstimatedRemainingTime { get; set; }

        public TaskProgress()
        {
            Percentage = 0;
            Message = string.Empty;
            CurrentOperation = string.Empty;
            CurrentStep = 0;
            TotalSteps = 0;
            ElapsedTime = TimeSpan.Zero;
            EstimatedRemainingTime = TimeSpan.Zero;
        }

        public TaskProgress(int percentage, string message = "")
        {
            Percentage = Math.Max(0, Math.Min(100, percentage));
            Message = message ?? string.Empty;
            CurrentOperation = string.Empty;
            CurrentStep = 0;
            TotalSteps = 0;
            ElapsedTime = TimeSpan.Zero;
            EstimatedRemainingTime = TimeSpan.Zero;
        }

        public TaskProgress(int currentStep, int totalSteps, string message = "")
        {
            CurrentStep = currentStep;
            TotalSteps = totalSteps;
            Percentage = totalSteps > 0 ? (int)((double)currentStep / totalSteps * 100) : 0;
            Message = message ?? string.Empty;
            CurrentOperation = string.Empty;
            ElapsedTime = TimeSpan.Zero;
            EstimatedRemainingTime = TimeSpan.Zero;
        }
    }

    #endregion
}