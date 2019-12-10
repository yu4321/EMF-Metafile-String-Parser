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
    public class EMF2StringParser
    {
        public bool IsLoaded { get; set; } = false;
        public bool isLoggingEnabled { get; set; } = false;
        private Graphics dummy = null;
        public Metafile metafileToParse { get; set; } = null;
        private EnumerateMetafileProc metafileDelegate = null;
        public List<string> ParsedExtTextOutWs { get; set; } = new List<string>();
        /// <summary>
        /// Warning: Parsing for DrawString has never been tested on this version.
        /// </summary>
        public List<string> ParsedDrawStrings { get; set; } = new List<string>();
        /// <summary>
        /// Warning: Parsing for EmfSmallTextOut is not works correctly on this version. 
        /// </summary>
        public List<char> ParsedSmallTextOuts { get; set; } = new List<char>();
        /// <summary>
        /// Stores log if parsing failed in traversing.
        /// </summary>
        public List<KeyValuePair<EmfPlusRecordType, byte[]>> ParseFailingLogs { get; set; } = new List<KeyValuePair<EmfPlusRecordType, byte[]>>();
        /// <summary>
        /// Initialize the parser. File load is needed.
        /// </summary>
        public EMF2StringParser()
        {
            metafileDelegate = new EnumerateMetafileProc(MetafileCallback);
            dummy = FromImage(new Bitmap(78, 78));
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
            metafileToParse = newEmf;
            IsLoaded = true;
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
        public void LoadMetaFile(string path)
        {
            metafileToParse = new Metafile(path);
            IsLoaded = true;
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
                    metafileToParse = new Metafile(stream);
                }
                IsLoaded = true;
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
        /// <param name="splPath">Path to SPL spool file</param>
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
        /// Start parsing the loaded data. Cannot be executed when EMF is not loaded.
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
            ParseFailingLogs.Clear();
            try
            {
                dummy.EnumerateMetafile(metafileToParse, new Point(0, 0), metafileDelegate);
                return true;
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
                    }
                    catch
                    {
                        if (isLoggingEnabled)
                        {
                            ParseFailingLogs.Add(new KeyValuePair<EmfPlusRecordType, byte[]>(recordType, dataArray));
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
                    }
                    catch
                    {
                        if (isLoggingEnabled)
                        {
                            ParseFailingLogs.Add(new KeyValuePair<EmfPlusRecordType, byte[]>(recordType, dataArray));
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
                        }
                    }
                    catch
                    {
                        if (isLoggingEnabled)
                        {
                            ParseFailingLogs.Add(new KeyValuePair<EmfPlusRecordType, byte[]>(recordType, dataArray));
                        }
                    }
                }
            }

            metafileToParse.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }
    }
}
