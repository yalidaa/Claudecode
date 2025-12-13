using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagneticField
{
    internal class CheckedListBoxItem
    {
        public string Name { get; set; }
        public Color LineColor { get; set; }

        public CheckedListBoxItem(string name, Color lineColor)
        {
            Name = name;
            LineColor = lineColor;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
