using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using CommandLine;

namespace docxGenerator4DataCLI
{
    public class DocxGeneratorConfig
    {
        [Option('t', "DocxTemplatePath", Required = true, HelpText = "Absolute or relative file path to *.DOCX template (with text anchors)")]
        public string? DocxTemplatePath { get; set; }

        [Option('o', "OutputDirectory", Required = true, HelpText = "Absolute or relative file path to output directory (where result docx files will be created)")]
        public string? OutputDirectory { get; set; }

        [Option('a', "AnchorFilePath", Required = true, HelpText = "Absolute or relative file path to table-file (one row represent one file, the column-values are the replaceable text for anchor's name in header-column)")]
        public string? AnchorFilePath { get; set; }

        [Option('s', "AnchorFileSeparator", Required = true, HelpText = "A string -- separator of columns in input table-file (by AnchorFilePath)")]
        public string? AnchorFileSeparator { get; set; }

        [Option('n', "NewFileName", Required = false, HelpText = "A name of createable DOCX file (default a file1, file2 ...). Use an anchor's name from AnchorFile to specify a need name from column; if file will exists with that name, it will be re-writed")]
        public string? NewFileName{ get; set; }

        [Option('r', "UseRelativePaths", Required = false, HelpText = "A flag, if true -- DocxTemplatePath, OutputDirectory, AnchorFilePath are reads as relative paths from directory with execution assembly. If false (by default), all paths are absolute")]
        public bool? UseRelativePaths { get; set; } = false;

        public void ConvertPaths()
        {
            if (UseRelativePaths != true) return;
            string execDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            if (DocxTemplatePath != null) DocxTemplatePath = Path.Combine(execDir, DocxTemplatePath);
            if (OutputDirectory != null) OutputDirectory = Path.Combine(execDir, OutputDirectory);
            if (AnchorFilePath != null) AnchorFilePath = Path.Combine(execDir, AnchorFilePath);

        }
    }
}
