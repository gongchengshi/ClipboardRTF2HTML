using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ClipboardRTF2HTML
{
   class Program
   { 
      [STAThread]
      static void Main(string[] args)
      {
         try
         {
            var rtf = System.Windows.Clipboard.GetText(System.Windows.TextDataFormat.Rtf);
         
            string html;
            try
            {
               html = RTF2HTML.RtfText2tHtml(rtf);

               var doc = new HtmlDocument();
               using (var reader = new StringReader(html))
               {
                  doc.Load(reader);

                  SetFragment(doc);
                  AdjustMarkup(doc);

                  using (var writer = new StringWriter())
                  {
                     doc.Save(writer);
                     html = writer.ToString();
                  }
               }
            }
            catch (Exception ex)
            {
               html = "<!--StartFragment-->\n<p>Failed to convert RTF to HTML</p>\n<pre>" +
                  ex + "</pre>\n<!--EndFragment-->";
               MessageBox.Show(ex.ToString());
            }

            System.Windows.Clipboard.SetText(html, System.Windows.TextDataFormat.Html);
         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.ToString());
            throw;
         }
      }

      static public void AdjustMarkup(HtmlDocument doc)
      {
         var pElements = doc.DocumentNode.Descendants()
            .Where(x => x.Name == "p");

         foreach(var p in pElements.ToList())
         {
            p.Name = "span";
            var br = doc.CreateElement("br");
            p.ParentNode.InsertAfter(br, p);
         }
      }

      static public void SetFragment(HtmlDocument doc)
      {
         var body = doc.DocumentNode.Descendants().FirstOrDefault(n => n.Name == "body");

         if (body != null)
         {
            var startFragment = HtmlCommentNode.CreateNode("<!--StartFragment-->");
            body.ChildNodes.Insert(0, startFragment);

            var endFragment = HtmlCommentNode.CreateNode("<!--EndFragment-->");
            body.ChildNodes.Append(endFragment);
         }
      }
   }
}
