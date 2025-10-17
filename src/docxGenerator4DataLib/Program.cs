using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CommandLine;

namespace docxGenerator4DataCLI
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var parser = new CommandLine.Parser(with => with.HelpWriter = null);
            var result = parser.ParseArguments<DocxGeneratorConfig>(args);

            result.WithParsed(options => DocxGeneratorFunc.Run(options))
                    .WithNotParsed(errors => DocxGeneratorFunc.DisplayHelp(result, errors));


            Console.WriteLine("\nEnd!");
#if RELEASE
            Console.ReadKey();
#endif
        }
    }
}
