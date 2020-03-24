# EMF2StringParser

A parser which can extract some texts from EMF Metafile and SPL spool file(which includes inner EMF file only). Early version yet.

# How To Use

  >### 1. Load Metafile which you want to parse, with LoadMetaFile(System.Drawing.Imaging.Metafile, string, and byte[] is supported) or LoadFromSPLFile(string, and byte[] is supported) or Constructor with parameter.
  
  >## Case A: Get Result as Lists of each RecordTypes
  
  >>### 2. Execute ParseStart().
    
  >>### 3. If Parsing was completed, results will be stored in ParsedExtTextOutWs, ParsedDrawStrings, ParsedSmallTextOuts as Generic.List .
  >## Case B: Get Result as one combined string
  
  >>### 2. Execute GetCombinedStringFromLoadedMetaFile();
    
  >>### 3. Result will be returned as one string. Linebreaks and Whitespaces between record will be identified by the elements of LineBreakCandidates and SpaceCandidates.
    
# Notes

  - This class is not tested properly and built for personal usage of contributor. Cannot guarantee the outcome of processed results.
    - Parse for SmallTextOut is not working properly now.
    - Parse for DrawString is not tested yet.
