using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace S2CLotusNotesExtractTool
{
    class Program
    {
        private static LotusNotesIntegrator ltsNtsIntegrator = null;
        static void Main(string[] args)
        {
            Console.WriteLine("Initializing the document extract process");
            RunExtractDocuments();
            Console.WriteLine("Finished document extract process");
            Console.ReadLine();
        }

        private static void RunExtractDocuments()
        {
            try
            {
                bool isServerData = true;
                ltsNtsIntegrator = new LotusNotesIntegrator("siamcconaghy", "Workshop/Jellinbah", isServerData);
                ltsNtsIntegrator.ExtractGraphicsCafeDocuments();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
