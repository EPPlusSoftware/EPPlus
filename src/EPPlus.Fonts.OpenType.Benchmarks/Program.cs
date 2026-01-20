using BenchmarkDotNet.Running;

namespace EPPlus.Fonts.OpenType.Benchmarks
{
    // Program.cs - Entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            // Kör alla benchmark-klasser i assembly
            BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        }
    }
}