using ScottPlot.Colormaps;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagneticField
{
    public class MagPowerWaveCalculator
    {
        //  特定参数应从文件中读取,或在高级设置中才有
        public int PointPreScd { get; set; } = 20;  // 每秒默认点数
        public int PointMax { get; set; } = 8192;  // 最大输出点数
        private double DemagnetizationDutyCycle { get; } = 0.5;  // 占空比（常量，固定为50%）

        // 退磁波形公共参数
        public int TimeTotal { get; set; } = 200;  // 总时间（秒）
        public double DemagnetizationAmplitude { get; set; } = 200;  // 振幅
        public double DemagnetizationDecayCoefficient { get; set; } = 0.06;  // 衰减系数,指数衰减取 0.06，线性衰减取 1 ;
        public int DemagnetizationWaveType { get; set; } = 0;  // 退磁类型，波形类型：0=指数正弦波，1=线性正弦波，2=指数矩形波，3=线性矩形波
        public double DemagnetizationPeriod { get; set; } = 2;  // 周期（秒）


        //时间间隔 s
        public double TimeInterval()
        {
            double timeInterval = 1.0 / PointPreScd;
            return timeInterval;
        }
        //最大长度
        public int GetTimeTotal()
        {
            int PointDelta = 1000 / PointPreScd;             /// 每个点的时间间隔  ms
            int ColLeng = PointPreScd * TimeTotal + 1;     /// 数组长度  ms
            return ColLeng;
        }

        /// <summary>
        /// 计算输出时间数组
        /// </summary>
        /// <returns>时间数组int</returns>
        public int[] GetTimeCol()
        {
            int PointDelta = 1000 / PointPreScd;             /// 每个点的时间间隔  ms
            int ColLeng = PointPreScd * TimeTotal + 1;     /// 数组长度  ms

            int[] TimeCol = new int[ColLeng];

            for (int i = 0; i < ColLeng; i++)
            {
                TimeCol[i] = i * PointDelta;
            }

            return TimeCol;
        }

        /// <summary>
        /// 计算输出时间数组,根据大华要求移动半周期一个点，退磁时间使用
        /// </summary>
        /// <returns>时间数组int</returns>
        public int[] GetTimeMoveCol()
        {
            int[] TimeCol2 = GetTimeCol();
            int T = Convert.ToInt32(DemagnetizationPeriod * 1000);   //周期转换为ms，2int

            for (int i = 0; i < TimeCol2.Length; i++)
            {
                if ((TimeCol2[i] % T) == (T / 2))       //半周期倍数
                {
                    TimeCol2[i - 1] = TimeCol2[i - 1] + 24;    //前一个点后移24ms
                }
            }
            return TimeCol2;
        }

        /// <summary>
        /// 计算退磁波形值
        /// </summary>
        /// <returns>波形值</returns>
        public double[] GetDemagnetizationWave()
        {
            double[] TimeCol3 = Array.ConvertAll(GetTimeMoveCol(), x => (double)x / 1000);  ///时间数组，并转换为double
            double Frequency = 2 * Math.PI / DemagnetizationPeriod;     ///计算频率
            double[] Wave = new double[TimeCol3.Length];        ///定义输出内容

            switch (DemagnetizationWaveType)
            {
                case 0: // 指数正弦波
                    for (int i = 0; i < TimeCol3.Length; i++)
                    {
                        Wave[i] = DemagnetizationAmplitude * Math.Exp(-DemagnetizationDecayCoefficient * TimeCol3[i]) * Math.Sin(Frequency * TimeCol3[i]);
                    }
                    break;
                case 1: // 线性正弦波
                    for (int i = 0; i < TimeCol3.Length; i++)
                    {
                        Wave[i] = DemagnetizationAmplitude * (1 - DemagnetizationDecayCoefficient * TimeCol3[i] / TimeTotal) * Math.Sin(Frequency * TimeCol3[i]);
                    }
                    break;
                case 2: // 指数矩形波
                    for (int i = 0; i < TimeCol3.Length; i++)
                    {
                        double rectValue = (TimeCol3[i] % DemagnetizationPeriod) < (DemagnetizationDutyCycle * DemagnetizationPeriod) ? 1.0 : -1.0;
                        Wave[i] = DemagnetizationAmplitude * Math.Exp(-DemagnetizationDecayCoefficient * TimeCol3[i]) * rectValue;
                    }
                    break;
                case 3: // 线性矩形波
                    for (int i = 0; i < TimeCol3.Length; i++)
                    {
                        double rectValue = (TimeCol3[i] % DemagnetizationPeriod) < (DemagnetizationDutyCycle * DemagnetizationPeriod) ? 1.0 : -1.0;
                        Wave[i] = DemagnetizationAmplitude * (1 - DemagnetizationDecayCoefficient * TimeCol3[i] / TimeTotal) * rectValue;
                    }
                    break;
                default:
                    break;
            }

            return Wave;
        }

    }

    public class MagPowerMagnetizationWaveCalculator
    //  充磁类
    {
        //  特定参数应从文件中读取,或在高级设置中才有
        public double MagBI { get; set; } = 18.4854;   // 磁场,电流转换 mT/A
        public int PointPreScd { get; set; } = 20;  // 每秒默认点数 /s
        public int PointMax { get; set; } = 8192;  // 最大输出点数
        private double MagStopTime { get; set; } = 2;    // 输出完毕后保持 0 的时间

        //  公共参数
        public double MagHoldTime { get; set; } = 1;  // 磁场保持时间 s
        public double MagMax { get; set; } = 1;  // 磁场最大值 mT
        public double MagGradient { get; set; } = 1;  // 磁场变化率 mT/s


        //时间间隔 s
        public double TimeInterval()
        {
            double timeInterval = 1.0 / PointPreScd;
            return timeInterval;
        }
        //最大长度
        public int GetTimeTotal()
        {
            double TimeTotal = MagMax / MagGradient + MagHoldTime + MagStopTime;    // 总时间，s
            int TimeTotal2 = Convert.ToInt32(Math.Ceiling(TimeTotal));     //  总时间，向上取整s          
            int ColLeng = PointPreScd * TimeTotal2 + 1;     /// 数组长度  ms,不能大于    PointMax
            return ColLeng;
        }


        /// <summary>
        /// 计算输出时间数组
        /// </summary>
        /// <returns>时间数组int</returns>
        public int[] GetTimeCol()
        {
            double TimeTotal = MagMax / MagGradient + MagHoldTime + MagStopTime;    // 总时间，s
            int TimeTotal2 = Convert.ToInt32(Math.Ceiling(TimeTotal));     //  总时间，向上取整s          
            int ColLeng = PointPreScd * TimeTotal2 + 1;     /// 数组长度  ms,不能大于    PointMax

            int[] TimeCol = new int[ColLeng];
            int PointDelta = 1000 / PointPreScd;             /// 每个点的时间间隔  ms

            for (int i = 0; i < ColLeng; i++)
            {
                TimeCol[i] = i * PointDelta;
            }

            return TimeCol;
        }

        /// <summary>
        /// 计算输出波形数组
        /// </summary>
        /// <returns>波形数组double</returns>
        public double[] GetMagnetizationWave()
        {
            int[] TimeCol2 = GetTimeCol();  //  获得时间数组
            double[] Wave = new double[TimeCol2.Length];        //  定义输出波形数组

            // 计算应输出的电流值
            for (int i = 0; i < TimeCol2.Length; i++)
            {
                //  按照ms时间，分三段取值,
                if (TimeCol2[i] < MagMax / MagGradient * 1000)
                { Wave[i] = TimeCol2[i] * MagGradient / 1000; }
                else if (TimeCol2[i] < ((MagMax / MagGradient + MagHoldTime) * 1000))
                { Wave[i] = MagMax; }
                else
                { Wave[i] = 0; }

                Wave[i] = Wave[i] / MagBI;      //  转换为电流值,A
            }

            return Wave;
        }
    }


    internal class WaveCal
    { 
        // 公共属性
        public double A { get; set; }      // 振幅
        public double k { get; set; }      // 衰减系数
        public double T { get; set; }      // 周期（秒）
        public double D { get; } = 0.5;    // 占空比（常量，固定为50%）

        // 构造函数初始化参数
        public WaveCal(double amplitude, double decayCoefficient, double period)
        {
            A = amplitude;
            k = decayCoefficient;
            T = period;
        }

        // 计算角频率 w = 2π/T
        private double AngularFrequency => 2 * Math.PI / T;

        // 正弦波相关函数
        public double SinLiner(double t) => A * (1 - k * t) * Math.Sin(AngularFrequency * t);
        public double SinExp(double t) => A * Math.Exp(-k * t) * Math.Sin(AngularFrequency * t);

        // 矩形波相关函数
        public double CalculateRectangleWave(double t) =>
            (t % T) < (D * T) ? 1.0 : -1.0;

        // 线性矩形波：y = A * (1 - kt) * R(t/T)
        public double RectLiner(double t)
        {
            ValidateParameters(t);
            double rectValue = CalculateRectangleWave(t);
            return A * (1 - k * t) * rectValue;
        }

        // 指数矩形波：y = A * e^{-kt} * R(t/T)
        public double RectExp(double t)
        {
            ValidateParameters(t);
            double rectValue = CalculateRectangleWave(t);
            return A * Math.Exp(-k * t) * rectValue;
        }

        // 验证参数有效性
        private void ValidateParameters(double t)
        {
            if (t < 0)
                throw new ArgumentException("时间t不能为负数", nameof(t));

            if (T <= 0)
                throw new InvalidOperationException("周期T必须大于0");
        }

        public double GetVoltage(int intType, double t)
        {
            double rectValue;
            if (intType == 1)
            {
                rectValue = SinLiner(t);
            }
            else if (intType == 2)
            {
                rectValue = RectExp(t);
            }
            else if (intType == 3)
            {
                rectValue = RectLiner(t);
            }
            else
            {
                rectValue = SinExp(t);
            }
            rectValue = Math.Round(rectValue, 3);
            return rectValue;
        }

    }
}
