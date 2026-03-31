namespace BindDataFromDiffSource.Models
{
    /// <summary>
    /// Represents a simple data point with Month and Value for charting
    /// </summary>
    public class DataPoint
    {
        public string Month { get; set; } = string.Empty;
        public double Value { get; set; }
    }
}
