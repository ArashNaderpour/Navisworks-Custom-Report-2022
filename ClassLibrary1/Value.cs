using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomReport2022
{
    class Value
    {
        public string Type { get; set; }
        public int Version { get; set; }
        public int Thickness { get; set; }
        public List<double> Color { get; set; }
        public List<double> MidPoint { get; set; }
        public List<double> MaxPoint { get; set; }
        public List<double> Origin { get; set; }
        public string Text { get; set; }
        public List<double> Start { get; set; }
        public List<double> @End { get; set; }
    }
}
