using System.Collections.Generic;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace docxGenerator4DataCLI
{
    public class DocxEditor
    {
        public DocxEditor(string docxPath)
        {
            if (!File.Exists(docxPath)) throw new FileNotFoundException(docxPath);
            mDoc = DocX.Load(docxPath);
        }

        public void ReplaceText(Dictionary<string, string> anchor2texts)
        {
            if (mDoc == null) return;
            foreach (var anchor2text in anchor2texts)
            {
                StringReplaceTextOptions sOpts = new StringReplaceTextOptions()
                {
                    SearchValue = anchor2text.Key,
                    NewValue = anchor2text.Value
                };

                mDoc.ReplaceText(sOpts);
            }
        }

        public void Close(bool isSave)
        {
            if (mDoc == null) return;

            if (isSave) mDoc.Save();
            mDoc.Dispose();
        }

        private string? mDocxPath;
        private DocX? mDoc;
    }
}
