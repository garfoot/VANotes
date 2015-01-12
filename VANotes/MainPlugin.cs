using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Net.Mime;
using System.Speech.Recognition;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace VANotes
{
    public class MainPlugin
    {
        private const string NotesFoundKey = "notesFound";
        private const string NoteResultKey = "note";
        private const string ChoiceIdKey = "noteChoices";
        private const string SearchTermKey = "noteSearchTerm";
        private const double ConfidenceLevel = 0.0;

        private static readonly Guid PluginId = new Guid("120A82D5-A747-446D-B5E7-6EC175B39B6D");
        private static OneNote.Application _oneNote;

        public static string VA_DisplayName()
        {
            return "VoiceAttack Notes plugin.";
        }

        public static string VA_DisplayInfo()
        {
            return "VoiceAttack Notes plugin.\r\n\r\nLets you take and search notes using VoiceAttack.";
        }

        public static Guid VA_Id()
        {
            return PluginId;
        }

        public static void VA_Init1(ref Dictionary<string, object> state, ref Dictionary<string, Int16?> conditions,
                                    ref Dictionary<string, string> textValues, ref Dictionary<string, object> extendedValues)
        {
            _oneNote = new OneNote.Application();
        }

        public static void VA_Exit1(ref Dictionary<string, object> state)
        {
        }

        /// <summary>
        ///     This function is where you will do all of your work. When VoiceAttack encounters an 'Execute External Plugin Function'
        ///     action, the plugin indicated will be called.
        /// </summary>
        /// <param name="context">
        ///     A string that can be anything you want it to be. This is passed in from the command action.
        ///     This was added to allow you to just pass a value into the plugin in a simple fashion
        ///     (without having to set conditions/text values beforehand). Convert the string to whatever type you need to.
        /// </param>
        /// <param name="state">
        ///     All values from the state maintained by VoiceAttack for this plugin. The state allows you to maintain
        ///     kind of a, 'session' within VoiceAttack. This value is not persisted to disk and will be erased on restart.
        ///     Other plugins do not have access to this state (private to the plugin).
        /// 
        ///     The state dictionary is the complete state. You can manipulate it however you want,
        ///     the whole thing will be copied back and replace what VoiceAttack is holding on to.
        /// </param>
        /// <param name="conditions">The conditions that were specified in the 'Execute External' command action.</param>
        /// <param name="textValues">The text values that were specified in the 'Execute External' command action.</param>
        /// <param name="extendedValues">Reserved, will be null.</param>
        public static void VA_Invoke1(string context, ref Dictionary<string, object> state,
            ref Dictionary<string, Int16?> conditions,
            ref Dictionary<string, string> textValues, ref Dictionary<string, object> extendedValues)
        {
            Clear(conditions, textValues);

            if (string.Compare(context, "SearchNote", StringComparison.OrdinalIgnoreCase) == 0)
            {
                SearchNote(state, conditions, textValues, extendedValues);
            }
            else if (string.Compare(context, "SelectChoice", StringComparison.OrdinalIgnoreCase) == 0)
            {
                SelectChoice(state, conditions, textValues, extendedValues);
            }
        }

        private static void SelectChoice(Dictionary<string, object> state, Dictionary<string, short?> conditions,
                                         Dictionary<string, string> textValues, Dictionary<string, object> extendedValues)
        {
            var choices = (Choice[]) state[ChoiceIdKey];

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

        private static void PlayBeep()
        {
            System.Media.SystemSounds.Beep.Play();
        }

        private static void SearchNote(Dictionary<string, object> state, Dictionary<string, short?> conditions,
                                    Dictionary<string, string> textValues, Dictionary<string, object> extendedValues)
        {

            string start = null;
            string searchTerm = GetDictation();
            string result;


            textValues[SearchTermKey] = searchTerm;

            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                conditions[NotesFoundKey] = -1;
                return;
            }

            _oneNote.FindPages(start, searchTerm, out result, false, false, OneNote.XMLSchema.xs2013);

            try
            {
                var doc = XDocument.Parse(result);
                var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");


                var foundIds = doc.Descendants(ns + "Section")
                    .Where(i => (string) i.Attribute("name") == "Elite")
                    .SelectMany(i => i.Elements(ns + "Page"))
                    .Select(i => new Choice{Id = (string) i.Attribute("ID"), Name = (string)i.Attribute("name")})
                    .ToArray();

                conditions[NotesFoundKey] = (short)foundIds.Length;

                if (foundIds.Length == 1)
                {
                    ReadNote(textValues, foundIds.First().Id);
                }
                else if (foundIds.Length > 1)
                {
                    ReadChoices(state, textValues, foundIds);
                }
            }
            catch (Exception e)
            {
                
            }
        }

        private static string GetDictation()
        {
            try
            {
                using (var engine = new SpeechRecognitionEngine(CultureInfo.CurrentUICulture))
                {
                    var grammar = new DictationGrammar {Name = "default dictation", Enabled = true};
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
            }
            catch (Exception e)
            {
                
            }

            return null;
        }

        private static void ReadChoices(Dictionary<string, object> state, Dictionary<string, string> textValues, IList<Choice> foundIds)
        {
            state[ChoiceIdKey] = foundIds;

            var choices = new StringBuilder();
            choices.AppendFormat("Found {0} pages\r\n", foundIds.Count);
            int i = 1;
            foreach (var choice in foundIds)
            {
                choices.AppendFormat("Item {0}\r\n {1}\r\n", i++, choice.Name);
            }

            textValues[NoteResultKey] = choices.ToString();
        }

        private static void ReadNote(Dictionary<string, string> textValues, string pageId)
        {
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");
            string pageXml;
            _oneNote.GetPageContent(pageId, out pageXml, OneNote.PageInfo.piBasic, OneNote.XMLSchema.xs2013);

            var page = XDocument.Parse(pageXml);
            var text = page.Descendants(ns + "T")
                .Select((i => (string) i))
                .Aggregate(new StringBuilder(), (sb, i) =>
                {
                    sb.AppendLine(i);
                    return sb;
                }).ToString();

            textValues[NoteResultKey] = text;
        }

        private static void Clear(Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            if (conditions.ContainsKey(NotesFoundKey))
                conditions.Remove(NotesFoundKey);

            if (textValues.ContainsKey(NoteResultKey))
                conditions.Remove(NoteResultKey);
        }
    }
}
