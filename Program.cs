using BootCamperMachineConfiguration.Controller;
using static System.Console;

namespace BootCamperMachineConfiguration
{
    public static class Program
    {
        static void Main(string[] args)
        {
            WriteLine("Please Enter Email Address");
            var machineConfiguration = new MachineConfigurationController(ReadLine());
            machineConfiguration.UpdateBootCamperMachineDetails();
        }
    }
}
