using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using static System.Drawing.Graphics;

namespace EMF2StringParser
{
    /// <summary>
    /// Parser Class for text tags of EMF Metafile.
    /// </summary>
    public class EMF2StringParser : IDisposable
    {
        private bool extractTextMode = false;
        /// <summary>
        /// Show parser is ready to be executed or not.
        /// </summary>
        public bool IsLoaded
        {
            get
            {
                return _loadedMetafile != null;
            }
        }
        /// <summary>
        /// Shows that logging featrue for failed record is enabled or disabled. 
        /// </summary>
        public bool IsParseFailedLoggingEnabled { get; set; } = false;
        private Graphics dummy = null;

        private Metafile _loadedMetafile = null;
        private EnumerateMetafileProc metafileDelegate = null;

        /// <summary>
        /// If records between text contains one of the members of this array, Linebreak(\n) will be added to ParsedExpectedText.
        /// </summary>
        public EmfPlusRecordType[] LineBreakCandidates { get; set; }
        /// <summary>
        /// If records between text contains all members of this array, Linebreak(\n) will be added to ParsedExpectedText.
        /// </summary>
        public EmfPlusRecordType[] MustLineBreak { get; set; }
        /// <summary>
        /// If records between text contains one of the members of this array, Space(" ") will be added to ParsedExpectedText.
        /// </summary>
        public EmfPlusRecordType[] SpaceCandidates { get; set; }
        /// <summary>
        /// If records between text contains all members of this array, Space(" ") will be added to ParsedExpectedText.
        /// </summary>
        public EmfPlusRecordType[] MustSpace { get; set; }

        private List<EmfPlusRecordType> betweenTextCommands = new List<EmfPlusRecordType>();

        private string parsedExpectedText= "";

        /// <summary>
        /// Parsed result of ExtTextOutW Records.
        /// </summary>
        public List<string> ParsedExtTextOutWs { get; private set; } = new List<string>();

        /// <summary>
        /// Parsed result of DrawString Records.
        /// </summary>
        public List<string> ParsedDrawStrings { get; private set; } = new List<string>();

        /// <summary>
        /// Parsed result of SmallTextOut Records.
        /// </summary>
        public List<char> ParsedSmallTextOuts { get; private set; } = new List<char>();

        /// <summary>
        /// Stores record if it failed to be parsed in traversing.
        /// </summary>
        public List<KeyValuePair<EmfPlusRecordType, byte[]>> ParseFailedRecords { get; private set; } = new List<KeyValuePair<EmfPlusRecordType, byte[]>>();
        /// <summary>
        /// Initialize the parser. File load is needed.
        /// </summary>
        public EMF2StringParser()
        {
            metafileDelegate = new EnumerateMetafileProc(MetafileCallback);
            dummy = FromImage(new Bitmap(78, 78));
        }

        /// <summary>
        /// Unload the loaded file.
        /// </summary>
        public void UnloadMetafile()
        {
            _loadedMetafile.Dispose();
            _loadedMetafile = null;
        }

        /// <summary>
        /// Initialize the parser. Automatically calls LoadMetaFile after initialized.
        /// </summary>
        /// <param name="newEmf">EMF file you which want to parse.</param>
        public EMF2StringParser(Metafile newEmf) : this()
        {
            LoadMetaFile(newEmf);
        }

        /// <summary>
        /// Ready for parse with given EMF file.
        /// </summary>
        /// <param name="newEmf">EMF file you which want to parse.</param>
        public void LoadMetaFile(Metafile newEmf)
        {
            _loadedMetafile = newEmf;
        }

        /// <summary>
        /// Initialize the parser. Automatically calls LoadMetaFile after initialized.
        /// </summary>
        /// <param name="emfPath">Relative path to EMF file you which want to parse.</param>
        public EMF2StringParser(string emfPath) : this()
        {
            LoadMetaFile(emfPath);
        }

        /// <summary>
        /// Ready for parse with given EMF file.
        /// </summary>
        /// <param name="path">Relative path to EMF file you which want to parse.</param>
        /// <param name="isCopyToMemory">If true, file will be loaded via copied by stream.</param>
        public void LoadMetaFile(string path, bool isCopyToMemory = false)
        {
            if (isCopyToMemory)
            {
                byte[] fileBytes = null;
                using (FileStream fs = File.Open(path, FileMode.Open))
                {
                    fileBytes = new byte[fs.Length];
                    fs.Read(fileBytes, 0, fileBytes.Length);
                    using (MemoryStream stream = new MemoryStream(fileBytes))
                    {
                        _loadedMetafile = new Metafile(stream);
                    }
                }
            }
            else
            {
                _loadedMetafile = new Metafile(path);
            }
        }

        /// <summary>
        /// Initialize the parser. Automatically calls LoadMetaFile after initialized.
        /// </summary>
        /// <param name="rawData">Byte array filled with data of EMF file.</param>
        public EMF2StringParser(byte[] rawData) : this()
        {
            LoadMetaFile(rawData);
        }

        /// <summary>
        /// Ready for parse with given EMF file.
        /// </summary>
        /// <param name="rawData">Byte array filled with data of EMF file.</param>
        public void LoadMetaFile(byte[] rawData)
        {
            try
            {
                using (MemoryStream stream = new MemoryStream(rawData))
                {
                    _loadedMetafile = new Metafile(stream);
                }
            }
            catch (Exception e)
            {
                throw new AggregateException(new Exception[] { e, new Exception("Make it sure that your EMF file contains EMF header correctly.") });
            }
        }


        /// <summary>
        /// Parse the SPL(Spool) File, Extract the EMF File, and LoadMetaFile with Extracted EMF File.
        /// </summary>
        /// <param name="splPath">Path to SPL spool file.</param>
        public void LoadFromSPLFile(string splPath)
        {
            LoadMetaFile(ExtractEMFfromSPL(splPath));
        }

        /// <summary>
        /// Parse the SPL(Spool) File, Extract the EMF File, and LoadMetaFile with Extracted EMF File.
        /// </summary>
        /// <param name="splFile">Byte Array of file contents.</param>
        public void LoadFromSPLFile(byte[] splFile)
        {
            LoadMetaFile(ExtractEMFfromSPL(splFile));
        }


        /// <summary>
        /// Extract the EMF File from SPL File
        /// </summary>
        /// <param name="splFile">Byte array of SPL File</param>
        /// <returns>Byte array of EMF File</returns>
        public static byte[] ExtractEMFfromSPL(byte[] splFile)
        {
            var position = BitConverter.ToInt32(splFile, 4);
            var size = BitConverter.ToInt32(splFile, position + 4);
            return splFile.Skip(position + 8).Take(size).ToArray();
        }

        /// <summary>
        /// Extract the EMF File from SPL File
        /// </summary>
        /// <param name="splPath">Path to SPL spool file</param>
        /// <returns>Byte array of EMF File</returns>
        public static byte[] ExtractEMFfromSPL(string splPath)
        {
            byte[] splFile = File.ReadAllBytes(splPath);
            return ExtractEMFfromSPL(splFile);
        }

        /// <summary>
        /// Start parsing the loaded data. Cannot be executed when EMF is not loaded. Every Parsed Datas will be stored in "Parsed(EMFRecordTypeName)".
        /// </summary>
        /// <returns></returns>
        public bool ParseStart()
        {
            if (IsLoaded == false)
            {
                throw new Exception("EMF File has not Loaded yet.");
            }
            ParsedExtTextOutWs.Clear();
            ParsedSmallTextOuts.Clear();
            ParsedDrawStrings.Clear();
            ParseFailedRecords.Clear();
            parsedExpectedText = "";
            try
            {
                dummy.EnumerateMetafile(_loadedMetafile, new Point(0, 0), metafileDelegate);
                return true;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Experimental : Start parsing and returns string. Every linebreak and space will be determined by Linebreak and Space array properties.
        /// </summary>
        /// <returns></returns>
        public string GetCombinedStringFromLoadedMetaFile()
        {
            if (IsLoaded == false)
            {
                throw new Exception("EMF File has not Loaded yet.");
            }
            ParsedExtTextOutWs.Clear();
            ParsedSmallTextOuts.Clear();
            ParsedDrawStrings.Clear();
            ParseFailedRecords.Clear();
            parsedExpectedText = "";
            extractTextMode = true;
            try
            {
                dummy.EnumerateMetafile(_loadedMetafile, new Point(0, 0), metafileDelegate);
                return parsedExpectedText;
            }
            catch
            {
                throw;
            }
        }

        private bool MetafileCallback(EmfPlusRecordType recordType, int flags, int dataSize, IntPtr data, PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;
            if (data != IntPtr.Zero)
            {
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);

                if (recordType == EmfPlusRecordType.DrawString)
                {
                    try
                    {
                        int stringLength = BitConverter.ToInt32(dataArray, 8) * 2;

                        string str = Encoding.Unicode.GetString(dataArray, 28, stringLength);
                        ParsedDrawStrings.Add(str);

                        if (extractTextMode)
                        {
                            GuessWhiteSpaces();

                            parsedExpectedText += str;
                        }
                    }
                    catch
                    {
                        if (IsParseFailedLoggingEnabled)
                        {
                            ParseFailedRecords.Add(new KeyValuePair<EmfPlusRecordType, byte[]>(recordType, dataArray));
                        }
                    }
                }
                else if (recordType == EmfPlusRecordType.EmfSmallTextOut)
                {
                    var x = BitConverter.ToInt32(dataArray, 0);
                    var y = BitConverter.ToInt32(dataArray, 4);
                    var cChars = BitConverter.ToInt32(dataArray, 8);
                    var fuOptions = BitConverter.ToInt32(dataArray, 12);
                    var iGraphicsMode = BitConverter.ToInt32(dataArray, 16);
                    var exScale = BitConverter.ToDouble(dataArray, 20);
                    var eyScale = BitConverter.ToDouble(dataArray, 24);
                    var maybeBound = BitConverter.ToInt32(dataArray, 28);
                    int num_chars = BitConverter.ToInt32(dataArray, 28);
                    try
                    {
                        var maybeText = BitConverter.ToChar(dataArray, 28);
                        ParsedSmallTextOuts.Add(maybeText);

                        if (extractTextMode)
                        {
                            GuessWhiteSpaces();

                            parsedExpectedText += maybeText;
                        }
                    }
                    catch
                    {
                        if (IsParseFailedLoggingEnabled)
                        {
                            ParseFailedRecords.Add(new KeyValuePair<EmfPlusRecordType, byte[]>(recordType, dataArray));
                        }
                    }
                }
                else if (recordType == EmfPlusRecordType.EmfExtTextOutW)
                {
                    string txt;
                    try
                    {
                        var length = BitConverter.ToUInt32(dataArray, 36);
                        var offString = BitConverter.ToUInt32(dataArray, 40);
                        var chars = new char[length];

                        for (int i = 0; i < length; i++)
                        {
                            chars[i] = BitConverter.ToChar(dataArray, (int)offString - 8 + i * 2);
                        }

                        txt = new string(chars);
                        if (txt.Replace(" ", "").Length > 0)
                        {
                            ParsedExtTextOutWs.Add(txt);

                            if (extractTextMode)
                            {
                                GuessWhiteSpaces();

                                parsedExpectedText += txt;
                            }
                        }
                    }
                    catch
                    {
                        if (IsParseFailedLoggingEnabled)
                        {
                            ParseFailedRecords.Add(new KeyValuePair<EmfPlusRecordType, byte[]>(recordType, dataArray));
                        }
                    }
                }
                else
                {
                    if (extractTextMode)
                    {
                        betweenTextCommands.Add(recordType);
                    }
                }

            }

            _loadedMetafile.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }

        private void GuessWhiteSpaces()
        {
            if (betweenTextCommands != null)
            {
                if (LineBreakCandidates != null && betweenTextCommands.Intersect(LineBreakCandidates).Count() > 0)
                {
                    parsedExpectedText += "\n";
                }
                else if (MustLineBreak != null && betweenTextCommands.Intersect(MustLineBreak).Count() == MustLineBreak.Count())
                {
                    parsedExpectedText += "\n";
                }
                else if (SpaceCandidates != null && betweenTextCommands.Intersect(SpaceCandidates).Count() > 0)
                {
                    parsedExpectedText += " ";
                }
                else if (MustSpace != null && betweenTextCommands.Intersect(MustSpace).Count() == MustSpace.Count())
                {
                    parsedExpectedText += " ";
                }
            }
            betweenTextCommands.Clear();
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    dummy.Dispose();
                    _loadedMetafile.Dispose();
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
