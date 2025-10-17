using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CommandLine.Text;
using CommandLine;

namespace docxGenerator4DataCLI
{
    public class DocxGeneratorFunc
    {
        public static void Run(DocxGeneratorConfig? config)
        {
            if (config == null) return;
            config.ConvertPaths();

            //Check input data
            if (!File.Exists(config.DocxTemplatePath ?? "")) throw new FileNotFoundException("docxGenerator4DataCLI. A docx template file is missing for path " + config.DocxTemplatePath ?? "");
            if (!File.Exists(config.AnchorFilePath ?? "")) throw new FileNotFoundException("docxGenerator4DataCLI. An anchor-file is missing for path " + config.AnchorFilePath ?? "");

            if (config.OutputDirectory == null) throw new Exception("docxGenerator4DataCLI. OutputDirectory is not specify");
            if (!Directory.Exists(config.OutputDirectory)) Directory.CreateDirectory(config.OutputDirectory);

            if (config.AnchorFileSeparator == null)
            {
                Console.WriteLine("docxGenerator4DataCLI. AnchorFileSeparator is not specify. The will using defaule value = TAB");
                config.AnchorFileSeparator = "\t";
            }

            // read anchor-file
            string[] anchorFileData = File.ReadAllLines(config.AnchorFilePath);

            string[] anchors = anchorFileData[0].Split(config.AnchorFileSeparator);

            int filesCounter = 1;
            foreach (string anchorFileString in anchorFileData.Skip(1))
            {
                if (!anchorFileString.Contains(config.AnchorFileSeparator)) continue;
                string[] anchorValues = anchorFileString.Split(config.AnchorFileSeparator);

                Dictionary<string, string> anchorData = new Dictionary<string, string>();
                for (int i = 0; i < anchors.Length; i++)
                {
                    anchorData[anchors[i]] = anchorValues[i];
                }

                string filePrefix = "file_" + filesCounter.ToString();
                if (config.NewFileName != null && anchors.Contains(config.NewFileName)) filePrefix = anchorData[config.NewFileName];
                string resultFilePath = Path.Combine(config.OutputDirectory, filePrefix + ".docx");
                if (File.Exists(resultFilePath))
                {
                    Console.WriteLine("docxGenerator4DataCLI. Result-file is exists, it will be re-write " + resultFilePath);
                }
                File.Copy(config.DocxTemplatePath, resultFilePath, true);

                DocxEditor editor = new DocxEditor(resultFilePath);
                editor.ReplaceText(anchorData);
                editor.Close(true);

                filesCounter++;
            }
        }

        public static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errors)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);

            Console.WriteLine(helpText);
        }
    }
}
