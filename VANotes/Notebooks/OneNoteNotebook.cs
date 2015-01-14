using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace VANotes.Notebooks
{
    public class OneNoteNotebook : INotebook
    {
        private readonly OneNote.Application _oneNote;

        public OneNoteNotebook()
        {
            _oneNote = new OneNote.Application();
            
        }

        public void ReadNote(VoiceAttack voiceAttack, string pageId)
        {
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");
            string pageXml;
            _oneNote.GetPageContent(pageId, out pageXml, OneNote.PageInfo.piBasic, OneNote.XMLSchema.xs2013);

            var page = XDocument.Parse(pageXml);
            var text = page.Descendants(ns + "T")
                .Select((i => (string)i))
                .Aggregate(new StringBuilder(), (sb, i) =>
                {
                    sb.AppendLine(i);
                    return sb;
                }).ToString();

            voiceAttack.Say(text);
        }

        public IList<Choice> FindPages(string searchTerm)
        {
            string result;
            _oneNote.FindPages(null, searchTerm, out result, false, false, OneNote.XMLSchema.xs2013);
        
            var doc = XDocument.Parse(result);
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");


            return doc.Descendants(ns + "Section")
                .Where(i => (string)i.Attribute("name") == "Elite")
                .SelectMany(i => i.Elements(ns + "Page"))
                .Select(i => new Choice { Id = (string)i.Attribute("ID"), Name = (string)i.Attribute("name") })
                .ToArray();
        }

        public string GetPage(string pageId)
        {
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");
            string pageXml;
            _oneNote.GetPageContent(pageId, out pageXml, OneNote.PageInfo.piBasic, OneNote.XMLSchema.xs2013);

            var page = XDocument.Parse(pageXml);
            var text = page.Descendants(ns + "T")
                .Select((i => (string)i))
                .Aggregate(new StringBuilder(), (sb, i) =>
                {
                    sb.AppendLine(i);
                    return sb;
                }).ToString();

            return text;
        }
    }
}