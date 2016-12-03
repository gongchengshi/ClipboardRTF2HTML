
namespace ClipboardRTF2HTML
{
   class Supplementary
   {
      static public string htmlClipboardHeader =
         "Version:1.0\r\n" +
         "StartHTML:{0:D9}\r\n" +
         "EndHTML:{1:D9}\r\n" +
         "StartFragment:{2:D9}\r\n" +
         "EndFragment:{3:D9}\r\n" +
         "SourceURL:file:///temp.rtf\r\n";

      static public string BasicHTML =
         "<!DOCTYPE><HTML><HEAD><TITLE> The HTML Clipboard</TITLE><BASE HREF=\"http://sample/specs\"></HEAD>" +
            "<BODY><UL><!--StartFragment--><LI> The Fragment </LI><!--EndFragment--></UL></BODY></HTML>";

      static public string BasicHTML2 =
         "<!DOCTYPE><HTML><HEAD><TITLE> The HTML Clipboard</TITLE><BASE HREF=\"http://sample/specs\"></HEAD>" +
            "<BODY><P>text2</p><P>text2</p><P>text2</p><P>text2</p></BODY></HTML>";


   }
}
