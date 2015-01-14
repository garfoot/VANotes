using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using System.Xml.Linq;

namespace VANotes.Commands
{
    public class SearchNoteCommand : IPluginCommand
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNote;
        private const double ConfidenceLevel = 0.0;
        
        
        public string CommandName { get { return "SearchNote"; } }
        
        public void Init(Dictionary<string, object> state, Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            _oneNote = new Microsoft.Office.Interop.OneNote.Application();
        }

        public void Terminate(Dictionary<string, object> state)
        {
        }

        public void Invoke(Dictionary<string, object> state, Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            string start = null;
            string searchTerm = GetDictation();
            string result;


            textValues[Keys.SearchTermKey] = searchTerm;

            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                conditions[Keys.NotesFoundKey] = -1;
                return;
            }

            _oneNote.FindPages(start, searchTerm, out result, false, false, Microsoft.Office.Interop.OneNote.XMLSchema.xs2013);

            var doc = XDocument.Parse(result);
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");


            var foundIds = doc.Descendants(ns + "Section")
                .Where(i => (string)i.Attribute("name") == "Elite")
                .SelectMany(i => i.Elements(ns + "Page"))
                .Select(i => new Choice { Id = (string)i.Attribute("ID"), Name = (string)i.Attribute("name") })
                .ToArray();

            conditions[Keys.NotesFoundKey] = (short)foundIds.Length;

            if (foundIds.Length == 1)
            {
                ReadNote(textValues, foundIds.First().Id);
            }
            else if (foundIds.Length > 1)
            {
                ReadChoices(state, textValues, foundIds);
            }
        }

        private static void ReadChoices(Dictionary<string, object> state, Dictionary<string, string> textValues, IList<Choice> foundIds)
        {
            state[Keys.ChoiceIdKey] = foundIds;

            var choices = new StringBuilder();
            choices.AppendFormat("Found {0} pages\r\n", foundIds.Count);
            int i = 1;
            foreach (var choice in foundIds)
            {
                choices.AppendFormat("Item {0}\r\n {1}\r\n", i++, choice.Name);
            }

            textValues[Keys.NoteResultKey] = choices.ToString();
        }

        private static string GetDictation()
        {
            using (var engine = new SpeechRecognitionEngine(CultureInfo.CurrentUICulture))
            {
                var grammar = new DictationGrammar { Name = "default dictation", Enabled = true };
                engine.LoadGrammar(grammar);
                grammar.SetDictationContext("Search for", null);

                engine.SetInputToDefaultAudioDevice();
                PlayBeep();
                var result = engine.Recognize();

                if (result != null && result.Confidence > ConfidenceLevel)
                {
                    return result.Text;
                }
            }

            return null;
        }

        private void ReadNote(Dictionary<string, string> textValues, string pageId)
        {
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");
            string pageXml;
            _oneNote.GetPageContent(pageId, out pageXml, Microsoft.Office.Interop.OneNote.PageInfo.piBasic, Microsoft.Office.Interop.OneNote.XMLSchema.xs2013);

            var page = XDocument.Parse(pageXml);
            var text = page.Descendants(ns + "T")
                .Select((i => (string)i))
                .Aggregate(new StringBuilder(), (sb, i) =>
                {
                    sb.AppendLine(i);
                    return sb;
                }).ToString();

            textValues[Keys.NoteResultKey] = text;
        }
        
        private static void PlayBeep()
        {
            System.Media.SystemSounds.Beep.Play();
        }
    }
}