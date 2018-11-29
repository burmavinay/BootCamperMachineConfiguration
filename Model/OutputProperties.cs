namespace BootCamperMachineConfiguration.Model
{
    public class OutputProperties
    {
        public string OperatingSystem { get; set; }
        public int MinimalCpuBenchmarkScore { get; set; }
        public double Memory { get; set; }
        public long Storage { get; set; }
        public long FreeDiskSpace { get; set; }
        public string Architecture { get; set; }
    }

}
