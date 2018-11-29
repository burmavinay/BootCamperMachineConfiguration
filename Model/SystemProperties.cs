
namespace BootCamperMachineConfiguration.Model
{
    public class SystemProperties
    {
        public string OperatingSystemName { get; set; }
        public string OperatingSystemEdition { get; set; }
        public string ProcessorName { get; set; }
        public int CpuBenchMarkScore { get; set; }
        public double UsableMemory { get; set; }
        public long StorageSpace { get; set; }
        public long FreeDiskSpace { get; set; }
        public string Architecture { get; set; }
    }

}
