using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Speech.Recognition;
using System.Text;
using System.Xml.Linq;

namespace VANotes.Commands
{
    public class SelectChoiceCommand : IPluginCommand
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNote;
        private const double ConfidenceLevel = 0.0;
        
        
        public string CommandName { get { return "SelectChoice"; }}

        public void Init(Dictionary<string, object> state, Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            _oneNote = new Microsoft.Office.Interop.OneNote.Application();
        }

        public void Terminate(Dictionary<string, object> state)
        {
        }

        public void Invoke(Dictionary<string, object> state, Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            var choices = (Choice[])state[Keys.ChoiceIdKey];

            var grammar = new GrammarBuilder();
            string[] choicesIndexes = Enumerable.Range(1, choices.Length).Select(i => i.ToString()).ToArray();

            grammar.Append(new Choices(choicesIndexes));

            using (var recognizer = new SpeechRecognitionEngine(CultureInfo.CurrentUICulture))
            {
                recognizer.LoadGrammar(new Grammar(grammar));
                recognizer.SetInputToDefaultAudioDevice();

                PlayBeep();
                var result = recognizer.Recognize();

                if (result != null && result.Confidence > ConfidenceLevel)
                {
                    string res = result.Text;
                    int index;
                    if (int.TryParse(res, out index))
                    {
                        ReadNote(textValues, choices[index - 1].Id);
                    }
                }
            }
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