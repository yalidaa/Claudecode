using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagneticField
{
    internal class Magnetization
    {

        // 公共属性
        public double T { get; set; }      // 时长
        public double B { get; set; }      // 磁场强度
        public double K { get; set; }      // 变化率
        public double BI { get; set; }     // BI常数

        // 构造函数初始化参数
        public Magnetization(double time, double power, double change, double constant)
        {
            T = time;
            B = power;
            K = change;
            BI = constant;
        }

        public double GetMagnetizingMagneticFieldStrength(double t)//Get充磁磁场强度
        {
            double b = 0;

            // 当 t > T 时，磁场强度为 0
            if (t > T)
            {
                b = 0;
            }
            // 当磁场强度达到最大值 B 时，维持不变
            else if (K * t >= B)
            {
                b = B;
            }
            // 磁场强度随时间线性增加
            else
            {
                b = K * t;
            }

            // 除以 BI 后返回真正的磁场强度值
            return b / BI;
        }

    }
}
