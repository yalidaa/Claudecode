using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagneticField
{
    //使用新建类代替List<double>，加快运算时间
    internal class CircularBuffer
    {
        private readonly double[] _buffer;
        private int _index = 0;
        private bool _isFull = false;
        private readonly object _lock = new object();
        private readonly int _capacity;

        public CircularBuffer(int capacity)
        {
            _capacity = capacity;
            _buffer = new double[capacity];
        }

        public void Add(double value)
        {
            lock (_lock)
            {
                _buffer[_index] = value;
                _index++;
                if (_index >= _capacity)
                {
                    _index = 0;
                    _isFull = true;
                }
            }
        }

        public double[] GetData()
        {
            lock (_lock)
            {
                if (!_isFull)
                {
                    // 缓冲区未满，返回实际数据
                    var result = new double[_index];
                    Array.Copy(_buffer, 0, result, 0, _index);
                    return result;
                }
                else
                {
                    // 缓冲区已满，重新排序数据（最新的数据在最后）
                    var result = new double[_buffer.Length];
                    Array.Copy(_buffer, _index, result, 0, _buffer.Length - _index);
                    Array.Copy(_buffer, 0, result, _buffer.Length - _index, _index);
                    return result;
                }
            }
        }

        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _isFull ? _buffer.Length : _index;
                }
            }
        }

        public int Capacity => _capacity;

        public void Clear()
        {
            lock (_lock)
            {
                _index = 0;
                _isFull = false;
                Array.Clear(_buffer, 0, _buffer.Length);
            }
        }

        //仅保留最后 keepCount 个元素，用于清图时的残留数据
        public void KeepLast(int keepCount)
        {
            if (keepCount < 0) throw new ArgumentException("keepCount must be >= 0");
            if (keepCount == 0)
            {
                Clear();
                return;
            }

            lock (_lock)
            {
                int currentCount = _isFull ? _capacity : _index;

                if (keepCount >= currentCount)
                {
                    // 无需操作，已满足
                    return;
                }

                // 获取当前所有数据（按时间顺序）
                double[] allData = GetData(); // 注意：GetData() 返回的是 [最早, ..., 最新]

                // 我们要保留最后 keepCount 个（即最新 keepCount 个）
                double[] keptData = new double[keepCount];
                Array.Copy(allData, allData.Length - keepCount, keptData, 0, keepCount);

                // 重置缓冲区，并重新写入保留的数据
                _index = 0;
                _isFull = false;
                Array.Clear(_buffer, 0, _buffer.Length);

                foreach (double val in keptData)
                {
                    _buffer[_index] = val;
                    _index++;
                    if (_index >= _capacity)
                    {
                        _index = 0;
                        _isFull = true;
                    }
                }
            }
        }

    }
}
