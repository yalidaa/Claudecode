using NationalInstruments.DAQmx;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace NationalInstruments.Examples.ContAcqVoltageSamples_IntClk
{
    public class AngleDetector : IDisposable
    {
        private AnalogMultiChannelReader analogInReader;
        private NationalInstruments.DAQmx.Task myTask;
        private NationalInstruments.DAQmx.Task runningTask;
        private AsyncCallback analogCallback;

        private AnalogWaveform<double>[] data;

        // 为线程安全访问添加锁对象
        private readonly object _processingLock = new object();
        private readonly object _stopLock = new object();

        // 新增：用于生产者-消费者模式的队列和任务
        private BlockingCollection<AnalogWaveform<double>[]> dataQueue;
        private System.Threading.Tasks.Task processingTask;
        private CancellationTokenSource processingCts;

        // 状态变量
        private volatile int currentAngle = 0;
        private int prevChannelIndex = 0; // 0:a0, 1:a1, 2:a2, 3:a3
        private int countBelow1 = 0;

        // 通道标志位
        private bool[] flag1Triggered = new bool[4];
        private bool[] flag2Triggered = new bool[4];
        private bool[] flag3Triggered = new bool[4];
        private double[] curdata = new double[4];

        //  黑 第一个 第二 第三 通道顺序
        private int[] anglesouce = new int[4] { 0, 2, 3, 1 };

        // 新增：用于判断逆序的变量
        private int consecutiveChannelCount = 1;
        private const int CONSECUTIVE_CHANNEL_THRESHOLD = 20; // 连续检测到多次相同通道则认为可能发生逆序
        private const int FAST_ROTATION_THRESHOLD = 5; // 当上一个通道的采样计数小于此阈值时，认为速度过快

        // 代表当前旋转方向的全局变量
        private volatile bool isCurrentlyClockwise = true; // true表示顺时针，false表示逆时针

        // 新增：预处理阶段的状态机和数据结构
        private enum ProcessingState
        {
            Stopped,
            Preprocessing,
            WaitingForZero,
            Running
        }
        private ProcessingState currentState = ProcessingState.Stopped;
        private List<int> preprocessingChannelSequence = new List<int>();
        private const int PREPROCESSING_SEQUENCE_LENGTH = 5; // 收集2个不同的通道来确定初始方向

        // 如果mode=1采用 count个点进行计数  mode=0就单个处理
        int mode = 1;
        int count = 20;
        //两个阈值，第一个是开始记录角度，第二个是开始记录点数
        double threshold11 = 1.0;
        double threshold12 = 1.0;
        volatile int startflag = 0;

        public event Action<int> AngleChanged;
        public event Action<string> ErrorOccurred;
        public event Action<int> PacketLossDetected;
        public event Action<int> ReportRotationSpeed;
        public event Action<bool> DirectionChanged;
        public event Action<int> StateChanged;
        public event Action<AnalogWaveform<double>[]> DataAcquired;

        private string physicalChannel;
        private double minVoltage;
        private double maxVoltage;
        private double rate;
        private int samplesPerChannel;
        private bool isDisposed = false;

      //  private string logFilePath = Path.Combine(
      //"D:\\航天五院\\角度\\航天五院项目\\生成的结果",
      //$"curmapindex_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

        public AngleDetector()
        {

        }

        /// <summary>
        /// 获取当前角度值。
        /// </summary>
        /// <returns>当前角度</returns>
        //public int GetCurrentAngle()
        //{
        //    return currentAngle;
        //}

        /// <summary>
        /// 获取当前旋转方向。
        /// </summary>
        /// <returns>如果为 true，则为顺时针；否则为逆时针。</returns>
        public bool IsClockwise()
        {
            return isCurrentlyClockwise;
        }

        public void Start(string physicalChannel, double minVoltage, double maxVoltage, double rate, int samplesPerChannel, DataTable dataTable, DataColumn[] dataColumn, DataGrid acquisitionDataGrid)
        {
            //物理通道.一个 "Dev1/ai0:3" 的值表示使用名为 "Dev1" 的设备上的模拟输入通道 0 到 3。
            this.physicalChannel = physicalChannel;

            //最小/最大电压。这两个参数定义了您期望测量的模拟信号的电压范围。
            this.minVoltage = minVoltage;
            this.maxVoltage = maxVoltage;

            //rate: 采样率。这个参数以样本/秒（Hz）为单位，它决定了 DAQ 设备从每个通道采集电压样本的频率。
            this.rate = rate;

            //samplesPerChannel: 每通道采样数。这个参数指定了在每一次数据读取操作中，从每个通道采集多少个样本。
            this.samplesPerChannel = samplesPerChannel;

            if (runningTask == null)
            {
                try
                {
                    // 初始化队列和消费者任务
                    dataQueue = new BlockingCollection<AnalogWaveform<double>[]>();
                    processingCts = new CancellationTokenSource();
                    processingTask = System.Threading.Tasks.Task.Run(() => ConsumeData(processingCts.Token));

                    InitializeDetection();
                    ResetState();
                    currentState = ProcessingState.Preprocessing;// 设置初始状态为预处理
                    StateChanged?.Invoke((int)currentState);

                    myTask = new NationalInstruments.DAQmx.Task();

                    myTask.AIChannels.CreateVoltageChannel(physicalChannel, "",
                        (AITerminalConfiguration)(-1), minVoltage,
                        maxVoltage, AIVoltageUnits.Volts);

                    myTask.Timing.ConfigureSampleClock("", rate,
                        SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, 1000);
                    myTask.Stream.Buffer.InputBufferSize = 1000000;

                    myTask.Control(TaskAction.Verify);

                    // Prepare the table for Data
                    InitializeDataTable(myTask.AIChannels, dataColumn, ref dataTable);
                    acquisitionDataGrid.DataSource = dataTable;

                    runningTask = myTask;
                    analogInReader = new AnalogMultiChannelReader(myTask.Stream);
                    analogCallback = new AsyncCallback(AnalogInCallback);

                    // Use SynchronizeCallbacks to specify that the object 
                    // marshals callbacks across threads appropriately.
                    analogInReader.SynchronizeCallbacks = false;

                    analogInReader.BeginReadWaveform(samplesPerChannel,
                        analogCallback, myTask);
                }
                catch (DaqException exception)
                {
                    //ErrorOccurred?.Invoke(exception.Message);
                    if (myTask != null)
                    {
                        myTask.Dispose();
                    }
                    runningTask = null;
                }
            }
        }

        public void Stop()
        {
            if (myTask != null)
            {
                runningTask = null;
                myTask.Dispose();
                myTask = null;
            }
            // 请求取消消费者任务
            processingCts?.Cancel();
            dataQueue?.CompleteAdding(); // 确保即使没有数据，GetConsumingEnumerable 也能退出

            try
            {
                // 等待任务结束（由于取消，这会很快完成）
                processingTask?.Wait();
            }
            catch (AggregateException ex)
            {
                // 当任务因取消而结束时，Wait() 会抛出 AggregateException，其中包含 OperationCanceledException
                // 我们只处理这种预期的异常，以防止程序崩溃
                ex.Handle(e => e is OperationCanceledException);
            }
            finally
            {
                processingCts?.Dispose();
                processingCts = null;
            }

            ResetState();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (isDisposed) return;

            if (disposing)
            {
                Stop();
            }

            isDisposed = true;
        }

        private void ProcessData(AnalogWaveform<double>[] acquiredData)
        {
            //// 如果在获取锁时，状态已经被重置为 Stopped，则直接放弃处理这批“陈旧”数据
            if (currentState == ProcessingState.Stopped)
            {
                return;
            }

            if (mode == 0)
            {
                //可以支持来多个数据流，分别处理
                for (int i = 0; i < acquiredData[0].Samples.Count; i++)
                {
                    int channelCount = acquiredData.Length;
                    for (int ch = 0; ch < channelCount; ch++)
                    {
                        ProcessChannel(ch, acquiredData[ch].Samples[i].Value);
                        //每次处理完一个通道，就要清空通道状态
                        InitializeDetection();
                    }
                }
            }
            else
            {
                //可以支持来多个数据流，分别处理
                for (int i = 0; i < acquiredData[0].Samples.Count; i++)
                {
                    int channelCount = acquiredData.Length;
                    //处理一组每个通道的数据，进行分别累加
                    for (int ch = 0; ch < channelCount; ch++)
                    {
                        curdata[ch] += acquiredData[ch].Samples[i].Value;
                    }
                    //累计20组之后，进行处理
                    if ((i + 1) % count == 0)
                    {
                        for (int ch = 0; ch < channelCount; ch++)
                        {
                            ProcessChannel(ch, curdata[ch] / count);
                            //每次处理完一个通道，就要清空通道状态
                            InitializeDetection();
                            curdata[ch] = 0;
                        }
                    }
                }
            }

            DataAcquired?.Invoke(acquiredData);
        }

        private void AnalogInCallback(IAsyncResult ar)
        {
            try
            {   
                if (runningTask != null && runningTask == ar.AsyncState)
                {
                    // 在回调线程中结束异步读取
                    //var acquiredData = analogInReader.EndReadWaveform(ar);
                    var acquiredData = analogInReader.EndMemoryOptimizedReadWaveform(ar);

                    // 启动一个新的任务来处理数据，完全不阻塞回调线程
                    if (!dataQueue.IsAddingCompleted)
                    {
                        dataQueue.Add(acquiredData);
                    }

                    // 立即开始下一次读取
                    analogInReader.BeginMemoryOptimizedReadWaveform(samplesPerChannel, analogCallback, myTask, acquiredData);
                    //analogInReader.BeginReadWaveform(samplesPerChannel, analogCallback, myTask);
                }
            }
            catch (DaqException exception)
            {
                if (!isDisposed)
                {
                    if (myTask != null)
                    {
                        ErrorOccurred?.Invoke(exception.Message);
                        myTask.Dispose();
                        runningTask = null;
                    }
                }
                else
                {
                    runningTask = null;
                    myTask?.Dispose();
                }
            }
        }

        // 消费者方法
        private void ConsumeData(CancellationToken token)
        {
            try
            {
                // 此循环会阻塞直到有数据可用，或队列被标记为完成，或收到取消请求
                foreach (var acquiredData in dataQueue.GetConsumingEnumerable(token))
                {
                    // 此处的锁仍然是必要的，以防 Stop 方法与数据处理发生竞争
                    lock (_processingLock)
                    {
                        // 检查状态，以防在等待锁时任务已被停止
                        if (currentState != ProcessingState.Stopped)
                        {
                            ProcessData(acquiredData);
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // 任务被取消是正常操作，捕获异常以安静地退出循环
                // 无需额外处理
            }
        }

        private void ProcessChannel(int channel, double value)
        {
            //int curmapIndex = mapindex(channel);
            int curmapIndex = channel;
            if (!flag1Triggered[curmapIndex] && value <= threshold11)
            {
                flag1Triggered[curmapIndex] = true;
                //Console.WriteLine($"通道{channel}触发Flag1,map 通道是 {curmapIndex}");

                // 根据当前状态执行不同逻辑
                switch (currentState)
                {
                    case ProcessingState.Preprocessing:
                        HandlePreprocessing(curmapIndex);
                        break;
                    case ProcessingState.WaitingForZero:
                        HandleWaitingForZero(curmapIndex);
                        break;
                    case ProcessingState.Running:
                        UpdateAngle(curmapIndex);
                        //Console.WriteLine($"当前角度{currentAngle}度\n");
                        break;
                }

                //保存curmapIndex值到变量并追加到文件
                //AppendCurMapIndexToFile(curmapIndex, channel);
            }

            if (flag1Triggered[curmapIndex] && !flag2Triggered[curmapIndex] && value <= threshold12)
            {
                flag2Triggered[curmapIndex] = true;
                countBelow1 = 0;
            }

            if (flag2Triggered[curmapIndex] && !flag3Triggered[curmapIndex] && value > threshold12)
            {
                flag3Triggered[curmapIndex] = true;
            }

            if (flag2Triggered[curmapIndex] && !flag3Triggered[curmapIndex] && value <= threshold12)
            {
                countBelow1++;
            }
        }

        // 预处理阶段判断旋转方向
        private void HandlePreprocessing(int curmapIndex)
        {
            // 收集不重复的通道序列
            if (preprocessingChannelSequence.Count == 0 || preprocessingChannelSequence.Last() != curmapIndex)
            {
                preprocessingChannelSequence.Add(curmapIndex);
                //Console.WriteLine($"预处理阶段: 收集到通道 {curmapIndex}，当前序列: [{string.Join(", ", preprocessingChannelSequence)}]");
            }

            // 收集到足够的数据后判断方向
            if (preprocessingChannelSequence.Count >= PREPROCESSING_SEQUENCE_LENGTH)
            {
                int clockwiseVotes = 0;
                int counterClockwiseVotes = 0;
                // 遍历序列，根据相邻通道的变化进行投票
                for (int i = 0; i < preprocessingChannelSequence.Count - 1; i++)
                {
                    int first = preprocessingChannelSequence[i];
                    int second = preprocessingChannelSequence[i + 1];

                    if (isClockwiseVote(first,second))
                    {
                        isCurrentlyClockwise = true;
                        clockwiseVotes++;
                    }
                    else
                    {
                        isCurrentlyClockwise = false;
                        counterClockwiseVotes++;
                    }
                }

                // 只有在获得明确的投票结果时才确定方向
                if (clockwiseVotes > counterClockwiseVotes)
                {
                    isCurrentlyClockwise = true;
                    TransitionToWaitingForZero(curmapIndex);
                    return;
                }
                else if (counterClockwiseVotes > clockwiseVotes)
                {
                    isCurrentlyClockwise = false;
                    TransitionToWaitingForZero(curmapIndex);
                    return;
                }
                else
                {
                    // 如果票数相等（例如，来回抖动），则移除序列的第一个元素，等待下一个通道数据进来再判断
                    preprocessingChannelSequence.RemoveAt(0);
                    //Console.WriteLine("预处理方向不明确，尝试继续收集数据...");
                }
            }
            if (curmapIndex == 0)
            {
                //Console.WriteLine("预处理方向不明确，但是到0点了，所以使用最近确定的方向");
                TransitionToWaitingForZero(curmapIndex);
            }
        }

        private bool isClockwiseVote(int first, int second)
        {
            switch (second)
            {
                case 0:
                    if (first == 2) return  true; // 顺
                    else if (first == 1) return  false; // 逆
                    break;
                case 1:
                    if (first == 0 || first == 3) return true; // 顺
                    else if (first == 2) return false; // 逆
                    break;
                case 2:
                    if (first == 1) return true; // 顺
                    else if (first == 3 || first == 0)  return false; // 逆
                    break;
                case 3:
                    if (first == 2) return true; // 顺
                    else if (first == 1) return false; // 逆
                    break;
            }
            return true; // 默认顺时针
        }

        private void TransitionToWaitingForZero(int curmapIndex)
        {
            DirectionChanged?.Invoke(isCurrentlyClockwise);
            currentState = ProcessingState.WaitingForZero;
            StateChanged?.Invoke((int)currentState);
            Console.WriteLine($"预处理完成. 判断方向为: {(isCurrentlyClockwise ? "顺时针" : "逆时针")}. 进入等待0度位置状态.{Environment.NewLine}");

            // 检查当前最后一个值是否是0，如果是，直接进入Running状态
            if (curmapIndex == 0)
            {
                HandleWaitingForZero(curmapIndex);
            }
        }

        // 新增：处理等待0度位置状态的逻辑
        private void HandleWaitingForZero(int curmapIndex)
        {
            //Console.WriteLine("等待0度位置...");
            if (curmapIndex == 0)
            {
                //Console.WriteLine("检测到0度位置. 开始正式计算角度.");
                currentState = ProcessingState.Running;
                StateChanged?.Invoke((int)currentState);
                currentAngle = 0;
                prevChannelIndex = 0;
                startflag = 1; // 设置开始标志
            }
        }

        private int mapindex(int channel)
        {
            for (int i = 0; i < 4; i++)
            {
                if (channel == anglesouce[i])
                {
                    return i;
                }
            }
            Console.WriteLine($"！！！！！！！！！！！！！！！！！通道{channel}触发error， anglesouce中找不到对应的通道！！！！！！！！！！！！！");
            return 0;
        }

        private void UpdateAngle(int detectedChannel)
        {
            //如果之前一个角度是0，且当前通道不是0，这种情况是刚开始启动的时候进行的判断
            if (prevChannelIndex == 0 && detectedChannel == 0)
            {
                //初始化，启动信号，角度不变化
                //Console.WriteLine("a0开始检测");
                startflag = 1;
            }
            if (startflag == 1)
            {
                // 转动过程中 0度位置特殊处理
                if (prevChannelIndex == detectedChannel)
                {
                    consecutiveChannelCount++; // 如果检测到的通道和上一个相同，增加计数
                    //Console.WriteLine($"连续检测到通道 {detectedChannel}, 次数: {consecutiveChannelCount}");
                    return;
                }

                int angleBeforeUpdate;
                angleBeforeUpdate = currentAngle;

                // 根据实际的序列模式计算期望的下一个通道索引
                int expectedNextClockwise = GetExpectedNextChannel(angleBeforeUpdate, isCurrentlyClockwise);
                int expectedNextCounterClockwise = GetExpectedNextChannel(angleBeforeUpdate, !isCurrentlyClockwise);

                // 检测是否发生跳跃（丢包）
                int skippedSteps = 0;

                // 判断旋转方向和跳跃步数
                if (detectedChannel == expectedNextClockwise)
                {
                    // 正常沿着原来的方向，没有跳跃
                    skippedSteps = 0;

                    ReportRotationSpeed?.Invoke(consecutiveChannelCount);
                    consecutiveChannelCount = 1; // 方向确认，重置计数器
                }
                else if (detectedChannel == expectedNextCounterClockwise)
                {
                    // 可能是正常逆序旋转，也可能是顺序旋转丢包
                    // 使用连续相同通道的计数值来判断
                    if (consecutiveChannelCount >= CONSECUTIVE_CHANNEL_THRESHOLD)
                    {
                        // 计数值超过阈值，认为是真正的逆序旋转
                        isCurrentlyClockwise = !isCurrentlyClockwise;
                        DirectionChanged?.Invoke(isCurrentlyClockwise);

                        skippedSteps = 0;
                        //Console.WriteLine($"连续检测到通道 {prevChannelIndex} {{{consecutiveChannelCount}}} 次 >= 阈值 {CONSECUTIVE_CHANNEL_THRESHOLD}，判断为真正的逆序旋转");
                    }
                    else
                    {
                        // 计数值未到阈值，认为是继续原来的方向旋转但发生了丢包
                        skippedSteps = CalculateSkippedSteps(prevChannelIndex, detectedChannel, angleBeforeUpdate, isCurrentlyClockwise);
                        //Console.WriteLine($"连续检测到通道 {prevChannelIndex} {{{consecutiveChannelCount}}} 次 <= 阈值 {CONSECUTIVE_CHANNEL_THRESHOLD}，判断为顺序旋转丢包，跳过{skippedSteps}步");
                    }

                    ReportRotationSpeed?.Invoke(consecutiveChannelCount);
                    consecutiveChannelCount = 1; // 方向确认，重置计数器
                }
                else
                {
                    // 发生了跳跃，计算跳跃的步数和方向
                    // 使用连续相同通道计数值辅助判断方向倾向
                    if (consecutiveChannelCount > CONSECUTIVE_CHANNEL_THRESHOLD)
                    {
                        // 计数值超过阈值，倾向于逆序
                        isCurrentlyClockwise = !isCurrentlyClockwise;
                        DirectionChanged?.Invoke(isCurrentlyClockwise);

                        skippedSteps = CalculateSkippedSteps(prevChannelIndex, detectedChannel, angleBeforeUpdate, isCurrentlyClockwise);
                        //Console.WriteLine($"连续检测到通道 {prevChannelIndex} {{{consecutiveChannelCount}}} 次 > 阈值 {CONSECUTIVE_CHANNEL_THRESHOLD}，倾向于逆序，跳过{skippedSteps}步");
                    }
                    else
                    {
                        // 计数值未到阈值，倾向于顺序
                        skippedSteps = CalculateSkippedSteps(prevChannelIndex, detectedChannel, angleBeforeUpdate, isCurrentlyClockwise);
                        //Console.WriteLine($"连续检测到通道 {prevChannelIndex} {{{consecutiveChannelCount}}} 次 <= 阈值 {CONSECUTIVE_CHANNEL_THRESHOLD}，倾向于顺序，跳过{skippedSteps}步");
                    }

                    ReportRotationSpeed?.Invoke(consecutiveChannelCount);
                    consecutiveChannelCount = 1; // 方向确认，重置计数器
                }

                int newAngle;
                // 更新角度
                if (isCurrentlyClockwise)
                {
                    if (skippedSteps > 0)
                    {
                        for (int i = 1; i <= skippedSteps; i++)
                        {
                            int skippedAngle = (angleBeforeUpdate + i) % 360;
                            if (skippedAngle % 10 == 0 && skippedAngle != 0)
                            {
                                PacketLossDetected?.Invoke(skippedAngle);
                                break;
                            }
                        }
                        //Console.WriteLine($"检测到顺时针跳跃，从通道{prevChannelIndex}跳到通道{detectedChannel}，补偿角度：{skippedSteps}度");
                    }
                    newAngle = (angleBeforeUpdate + 1 + skippedSteps) % 360;
                }
                else
                {
                    if (skippedSteps > 0)
                    {
                        for (int i = 1; i <= skippedSteps; i++)
                        {
                            int skippedAngle = (angleBeforeUpdate - i + 360) % 360;
                            if (skippedAngle % 10 == 0 && skippedAngle != 0)
                            {
                                PacketLossDetected?.Invoke(skippedAngle);
                                break;
                            }
                        }
                        //Console.WriteLine($"检测到逆时针跳跃，从通道{prevChannelIndex}跳到通道{detectedChannel}，补偿角度：{skippedSteps}度");
                    }
                    newAngle = (angleBeforeUpdate - 1 - skippedSteps + 360) % 360;
                }

                currentAngle = newAngle;
                AngleChanged?.Invoke(newAngle);
                prevChannelIndex = detectedChannel;
            }
        }

        // 根据当前角度和通道索引，计算期望的下一个通道
        private int GetExpectedNextChannel(int currentAngle, bool clockwise)
        {
            int nextAngle;
            if (clockwise)
            {
                nextAngle = (currentAngle + 1) % 360;
            }
            else
            {
                nextAngle = (currentAngle - 1 + 360) % 360;
            }

            // 根据角度计算应该对应的通道索引
            return GetChannelIndexFromAngle(nextAngle);
        }

        // 根据角度获取对应的通道索引
        private int GetChannelIndexFromAngle(int angle)
        {
            // 序列模式：0123123123...123120 (总共360个位置)
            if (angle == 0)
            {
                return 0; // 黑色，0度位置
            }
            else
            {
                // 1-359度的位置按照123循环
                int position = (angle - 1) % 3; // 0, 1, 2 对应 1, 2, 3
                return position + 1; // 返回 1, 2, 3
            }
        }

        // 计算从当前通道跳跃到目标通道需要跳过的步数
        private int CalculateSkippedSteps(int fromChannel, int toChannel, int currentAngle, bool clockwise)
        {
            int steps = 0;
            int testAngle = currentAngle;
            int expectedChannel = fromChannel;

            // 最多检查360步，避免无限循环
            while (steps < 360 && expectedChannel != toChannel)
            {
                steps++;
                if (clockwise)
                {
                    testAngle = (testAngle + 1) % 360;
                }
                else
                {
                    testAngle = (testAngle - 1 + 360) % 360;
                }
                expectedChannel = GetChannelIndexFromAngle(testAngle);
            }

            return steps - 1; // 减1因为正常情况应该是1步
        }

        private void InitializeDetection()
        {
            countBelow1 = 0;
            for (int i = 0; i < 4; i++)
            {
                flag1Triggered[i] = false;
                flag2Triggered[i] = false;
                flag3Triggered[i] = false;
            }
        }

        public void InitializeDataTable(AIChannelCollection channelCollection, DataColumn[] dataColumn, ref DataTable data)
        {
            int numOfChannels = channelCollection.Count;
            data.Rows.Clear();
            data.Columns.Clear();
            dataColumn = new DataColumn[numOfChannels];
            int numOfRows = 10;

            for (int currentChannelIndex = 0; currentChannelIndex < numOfChannels; currentChannelIndex++)
            {
                dataColumn[currentChannelIndex] = new DataColumn();
                dataColumn[currentChannelIndex].DataType = typeof(double);
                dataColumn[currentChannelIndex].ColumnName = channelCollection[currentChannelIndex].PhysicalName;
            }

            data.Columns.AddRange(dataColumn);

            for (int currentDataIndex = 0; currentDataIndex < numOfRows; currentDataIndex++)
            {
                object[] rowArr = new object[numOfChannels];
                data.Rows.Add(rowArr);
            }
        }

        private void ResetState()
        {
            currentAngle = 0;
            prevChannelIndex = 0;
            countBelow1 = 0;
            startflag = 0;

            // 重置状态机
            currentState = ProcessingState.Stopped;
            StateChanged?.Invoke((int)currentState);

            preprocessingChannelSequence.Clear();

            // 重置逆序判断变量
            consecutiveChannelCount = 1;

            // 重置方向跟踪变量
            isCurrentlyClockwise = true;

            Array.Clear(flag1Triggered, 0, flag1Triggered.Length);
            Array.Clear(flag2Triggered, 0, flag2Triggered.Length);
            Array.Clear(flag3Triggered, 0, flag3Triggered.Length);
            Array.Clear(curdata, 0, curdata.Length);
            AngleChanged?.Invoke(0);
        }

        public void testfix()
        {
            currentState = ProcessingState.Preprocessing;

            ////不发生丢包，只是不停逆序摇摆,停留4次以上表示停止准备逆序了，请将CONSECUTIVE_CHANNEL_THRESHOLD改为4
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[0]);
            dealsingle(anglesouce[0]);
            dealsingle(anglesouce[0]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[3]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[3]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[3]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[3]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[3]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[3]);
            dealsingle(anglesouce[1]);
            dealsingle(anglesouce[2]);
            dealsingle(anglesouce[1]);
        }

        //模拟给每个通道信号
        private void dealsingle(int i)
        {
            ProcessChannel(i, 10.0); // 模拟信号开始，电压较高
            ProcessChannel(i, 2.0);  // 电压下降
            ProcessChannel(i, 0.8);  // 电压低于阈值，触发检测逻辑
            ProcessChannel(i, 0.1);  // 电压继续下降
            ProcessChannel(i, 3.0);  // 电压回升，完成一次脉冲
            InitializeDetection();   // 重置检测标志位，为下一次事件做准备
        }

        // 新增：将curmapIndex值追加到文件的方法
        //private void AppendCurMapIndexToFile(int curmapIndex, int channel)
        //{
        //    try
        //    {
        //        DateTime timestamp = DateTime.Now;
        //        string logEntry = $"{timestamp:yyyy-MM-dd HH:mm:ss.fff} - Channel: {channel}, CurMapIndex: {curmapIndex}, CurrentAngle: {currentAngle}{Environment.NewLine}";
        //        //File.AppendAllText(logFilePath, logEntry);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"写入文件失败: {ex.Message}");
        //    }
        //}
    }
}
