using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EMFTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var parser = new EMF2StringParser.EMF2StringParser();
            parser.LineBreakCandidates = new[] { EmfPlusRecordType.EmfSelectObject, EmfPlusRecordType.EmfDeleteObject };
            parser.SpaceCandidates = new[] { EmfPlusRecordType.EmfIntersectClipRect};
            var files = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.emf");
            if (files.Count() > 0)
            {
                foreach(var x in files)
                {
                    parser.LoadMetaFile(x);
                    parser.IsParseFailedLoggingEnabled = true;
                    var expected = parser.GetCombinedStringFromLoadedMetaFile();
                    if (expected != null)
                    {
                        File.WriteAllText($"{x.Replace(".emf", "")}_ExpectedTest.txt", expected);
                    }
                }
                MessageBox.Show(".emf file parse success.");
            }
            else
            {
                MessageBox.Show(".emf file not found.");
            }
        }
    }
}
