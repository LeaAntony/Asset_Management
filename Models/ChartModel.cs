using System.Data;

namespace Asset_Management.Models
{
    public class ChartModel
    {
        public List<string> labels { get; set; }
        public List<double> values { get; set; }
        public List<double> values2 { get; set; }
        public string label { get; set; }
        public double value1 { get; set; }
        public double value2 { get; set; }
        public ChartModel()
        {
            labels = new List<string>();
            values = new List<double>();
            values2 = new List<double>();
        }
    }
}
