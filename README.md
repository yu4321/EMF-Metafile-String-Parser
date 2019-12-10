# EMF2StringParser

A parser which can extract some texts from EMF Metafile and SPL spool file(which includes EMF file only). Early version yet.

# How To Use

  1. Load Metafile which you want to parse, with LoadMetaFile(System.Drawing.Imaging.Metafile, string, and byte[] is supported) or LoadFromSPLFile(string, and byte[] is supported) or Constructor with parameter.
  2. Execute ParseStart().
  3. If Parsing was completed, results will be stored in ParsedExtTextOutWs, ParsedDrawStrings, ParsedSmallTextOuts as Generic.List .
  
# Notes

  - This class is not tested properly and built for personal usage of contributor. Cannot guarantee the outcome of processed results.
    - Parse for SmallTextOut is not working properly now.
    - Parse for DrawString is not tested yet.
