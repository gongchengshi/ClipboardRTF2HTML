using System;
using System.IO;
using System.Net;
using Word = Microsoft.Office.Interop.Word;

namespace ClipboardRTF2HTML
{
   public class RTF2HTML
   {
      static public string RtfFile2tHtml(string rtfFile)
      {
         var applicationObject = new Word.Application();

         object missing = Type.Missing;
         object fileName = rtfFile;
         object False = false;
         applicationObject.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

         Word.Document documentObject =
            applicationObject.Documents.Open(ref fileName, ref missing, ref missing, ref missing,
                  ref missing, ref missing, ref missing, ref missing, ref missing,
                  ref missing, ref missing, ref False, ref missing, ref missing,
                   ref missing, ref missing);

         object tempFileName = Path.Combine(Path.GetTempPath(), "tempHtml.html");
         object fileFormt = Word.WdSaveFormat.wdFormatHTML;
         object makeFalse = false;
         object makeTrue = true;

         documentObject.SaveAs(ref tempFileName, ref fileFormt,
            ref makeFalse, ref missing, ref makeFalse,
            ref missing, ref missing, ref missing, ref makeFalse, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing);

         GC.Collect();
         GC.WaitForPendingFinalizers();
         documentObject.Close(ref makeFalse, ref missing, ref missing);
         GC.Collect();
         GC.WaitForPendingFinalizers();

         string html;

         //Extract html source from the temporary html file.

         var tempHtmlFile = tempFileName.ToString();

         var client = new WebClient();
         html = client.DownloadString(tempHtmlFile);
         GC.Collect();
         GC.WaitForPendingFinalizers();

         //using (var htmlFile = new StreamReader(tempHtmlFile))
         //{
         //   html = htmlFile.ReadToEnd();
         //}

         try
         {
            File.Delete(tempHtmlFile);
         }
         catch { }

         return html;
      }

      static public string RtfText2tHtml(string rtfText)
      {
         var tempRtfFilename = Path.Combine(Path.GetTempPath(), "tempRtf.rtf");

         File.Delete(tempRtfFilename);

         using (var file = new StreamWriter(tempRtfFilename))
         {
            file.Write(rtfText);
         }

         return RtfFile2tHtml(tempRtfFilename);
      }
   }
}
