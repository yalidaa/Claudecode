using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;


namespace MagneticField
{
    internal class UnifiedDataRecorder
    {
        private readonly Dictionary<ushort, Dictionary<string, DataRecord>> _buffer = new Dictionary<ushort, Dictionary<string, DataRecord>>();
        private readonly Dictionary<string, int> _probeNameToIndex;
        private readonly string _filePath;
        private readonly object _lock = new object();
        private bool _disposed = false;
        private readonly int _totalProbeCount;
        private readonly int _maxBufferedFrames = 50000;      // 最多缓存 1000 帧未完成数据
        private readonly ushort _maxSequenceGap = 10;       // 序列号跳变超过 10 视为旧数据失效
                                                            // --- 新增：写入队列和后台任务 ---
        private readonly BlockingCollection<string> _writeQueue = new BlockingCollection<string>(new ConcurrentQueue<string>());
        private readonly Task _writeTask;
        private StreamWriter _writer;

        public UnifiedDataRecorder(string folderPath, IEnumerable<ProbeInfo> allProbes)
        {
            if (allProbes == null || !allProbes.Any())
                throw new ArgumentException("探头列表不能为空");

            _probeNameToIndex = allProbes.ToDictionary(p => p.ProbeName, p => ExtractIndex(p.ProbeName));
            _totalProbeCount = _probeNameToIndex.Count;

            var sortedNames = allProbes.OrderBy(p => _probeNameToIndex[p.ProbeName])
                                       .Select(p => p.ProbeName)
                                       .ToList();

            var header = new List<string> { "Sequence" };
            foreach (var name in sortedNames)
            {
                int idx = _probeNameToIndex[name];
                header.Add($"c{idx}X");
                header.Add($"c{idx}Y");
                header.Add($"c{idx}Z");
            }

            string fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
            _filePath = Path.Combine(folderPath, fileName);

            Directory.CreateDirectory(folderPath);
            //File.WriteAllText(_filePath, string.Join(",", header) + Environment.NewLine);

            // 初始化 Writer，注意：去掉了 AutoFlush = true，利用缓冲区提高性能
            _writer = new StreamWriter(_filePath, true);

            // 写入表头
            _writer.WriteLine(string.Join(",", header));

            // 启动后台写入线程
            _writeTask = Task.Run(WriteLoop);
        }

        private static int ExtractIndex(string probeName)
        {
            var match = Regex.Match(probeName, @"探头(\d+)");
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }

        public void SubmitData(string probeName, DataRecord record)
        {
            if (_disposed) return;
            Dictionary<string, DataRecord> completedFrame = null;

            lock (_lock)
            {
                // 安全清理：如果 buffer 太大 或 sequence 跳变太大，清理旧数据
                CleanUpOldSequences(record.Sequence);

                if (!_buffer.TryGetValue(record.Sequence, out var dict))
                {
                    dict = new Dictionary<string, DataRecord>();
                    _buffer[record.Sequence] = dict;
                }

                if (!dict.ContainsKey(probeName))
                {
                    dict[probeName] = record;
                }

                // 如果收齐，立即写入并移除
                if (dict.Count == _totalProbeCount)
                {
                    //WriteLine(record.Sequence, dict);
                    // 凑齐一帧后，不再直接写入文件，而是加入队列
                    completedFrame = dict;
                    _buffer.Remove(record.Sequence);
                }
            }
            // --- 关键修改：在锁外部进行字符串格式化和入队 ---
            // 这样可以大幅减少锁的持有时间，提高并发吞吐量
            if (completedFrame != null)
            {
                EnqueueLine(record.Sequence, completedFrame);
            }
        }

        private void CleanUpOldSequences(ushort currentSeq)
        {
            if (_buffer.Count == 0) return;

            // 策略1：缓冲区太大 → 清理最旧的一半
            if (_buffer.Count > _maxBufferedFrames)
            {
                var keysToRemove = _buffer.Keys.OrderBy(k => k).Take(_maxBufferedFrames / 2).ToList();
                foreach (var key in keysToRemove)
                {
                    _buffer.Remove(key);
                }
                Console.WriteLine($"[UnifiedRecorder] 缓冲区超限，已清理 {_maxBufferedFrames / 2} 帧旧数据");
                return;
            }

            // 策略2：当前 seq 比最老的 seq 大太多 → 清理所有比 (currentSeq - maxGap) 小的
            //ushort minSeq = _buffer.Keys.Min();
            //if ((short)(currentSeq - minSeq) > _maxSequenceGap) // 注意 ushort 溢出，用 short 判断差值
            //{
            //    var keysToRemove = _buffer.Keys.Where(seq => (short)(currentSeq - seq) > _maxSequenceGap).ToList();
            //    foreach (var key in keysToRemove)
            //    {
            //        _buffer.Remove(key);
            //    }
            //    Console.WriteLine($"[UnifiedRecorder] 检测到序列跳变，清理 {keysToRemove.Count} 帧过期数据");
            //}
        }

        //private void WriteLine(ushort seq, Dictionary<string, DataRecord> data)
        //{
        //    var fields = new List<string> { seq.ToString() };

        //    var orderedProbes = _probeNameToIndex.OrderBy(kvp => kvp.Value).ToList();
        //    foreach (var kvp in orderedProbes)
        //    {
        //        string name = kvp.Key;
        //        if (data.TryGetValue(name, out var rec))
        //        {
        //            fields.Add(rec.X.ToString("F6"));
        //            fields.Add(rec.Y.ToString("F6"));
        //            fields.Add(rec.Z.ToString("F6"));
        //        }
        //        else
        //        {
        //            fields.Add(""); fields.Add(""); fields.Add("");
        //        }
        //    }

        //    try
        //    {
        //        //File.AppendAllText(_filePath, string.Join(",", fields) + Environment.NewLine);
        //        // 使用 StreamWriter 写入
        //        _writer.WriteLine(string.Join(",", fields));
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"写入失败: {ex.Message}");
        //    }
        //}

        // 将数据格式化为字符串并放入队列（极快，不涉及IO）
        private void EnqueueLine(ushort seq, Dictionary<string, DataRecord> data)
        {
            var fields = new List<string> { seq.ToString() };

            var orderedProbes = _probeNameToIndex.OrderBy(kvp => kvp.Value).ToList();
            foreach (var kvp in orderedProbes)
            {
                string name = kvp.Key;
                if (data.TryGetValue(name, out var rec))
                {
                    fields.Add(rec.X.ToString("F6"));
                    fields.Add(rec.Y.ToString("F6"));
                    fields.Add(rec.Z.ToString("F6"));
                }
                else
                {
                    fields.Add(""); fields.Add(""); fields.Add("");
                }
            }

            try
            {
                if (!_writeQueue.IsAddingCompleted)
                {
                    _writeQueue.Add(string.Join(",", fields));
                }
            }
            catch (ObjectDisposedException)
            {
                // 关键修复：忽略“集合已被释放”的异常
                // 这通常发生在停止记录的瞬间
            }
            catch (InvalidOperationException)
            {
                // 忽略“集合已标记为完成”的异常
            }
        }

        // 后台写入循环
        private void WriteLoop()
        {
            try
            {
                // GetConsumingEnumerable 会阻塞等待新数据，直到 CompleteAdding 被调用
                foreach (var line in _writeQueue.GetConsumingEnumerable())
                {
                    _writer.WriteLine(line);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"后台写入异常: {ex.Message}");
            }
            finally
            {
                // 关键修复：加 try-catch 保护 Flush，防止在 Dispose 后调用崩溃
                try
                {
                    _writer?.Flush();
                }
                catch { }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            // 1. 标记不再接受新数据
            _writeQueue.CompleteAdding();

            // 2. 等待后台写入任务完成（把队列里剩余的数据写完）
            try
            {
                _writeTask.Wait(2000); // 最多等待2秒
            }
            catch { }

            // 3. 关闭资源
            lock (_lock)
            {
                _writer?.Dispose();
                _writer = null;
                _writeQueue?.Dispose();
            }
            // 可选：停止时是否写出残缺帧？
            /*
            lock (_lock)
            {
                var sequences = _buffer.Keys.OrderBy(s => s).ToList();
                foreach (var seq in sequences)
                {
                    if (_buffer.TryGetValue(seq, out var dict))
                    {
                        WriteLine(seq, dict);
                    }
                }
                _buffer.Clear();
            }
            */
        }
    }


    public interface IDataSink
    {
        void OnDataReceived(string probeName, DataRecord record);
    }
}
