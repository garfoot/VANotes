using System;
using System.Collections.Generic;
using System.Globalization;
using System.Media;
using System.Speech.Recognition;

namespace VANotes
{
    public class VoiceAttack
    {
        private const double ConfidenceLevel = 0.0;
        private Dictionary<string, object> _state;
        private Dictionary<string, string> _textValues;
        private Dictionary<string, short?> _conditions;

        private static class TextKeys
        {
            public const string SayText = "VANotes.SayText";
            public const string SearchTerm = "VANotes.SearchTerm";
        }

        private static class ConditionKeys
        {
            public const string NotesFoundCount = "VANotes.NotesFoundCount";
            public const string ErrorCode = "VANotes.ErrorCode";
        }

        private static class StateKeys
        {
            public const string Choices = "VANotes.Choices";
        }


        public VoiceAttack(Dictionary<string, object> state, Dictionary<string, string> textValues, Dictionary<string, short?> conditions)
        {
            _state = state;
            _textValues = textValues;
            _conditions = conditions;
        }

        public IList<Choice> Choices
        {
            get
            {
                if(_state.ContainsKey(StateKeys.Choices))
                    return _state[StateKeys.Choices] as IList<Choice>;

                return null;
            }
            set
            {
                _state[StateKeys.Choices] = value;
            }
        }

        public string SearchTerm
        {
            get
            {
                if(_textValues.ContainsKey(TextKeys.SearchTerm))
                    return _textValues[TextKeys.SearchTerm];

                return null;
            }
            private set
            {
                _textValues[TextKeys.SearchTerm]  = value;
            }
        }

        public void Say(string text)
        {
            _textValues[TextKeys.SayText] = text;
        }

        public void Error(int errorCode)
        {
            _conditions[ConditionKeys.ErrorCode] = (short)errorCode;
        }

        public void FoundPages(int count)
        {
            _conditions[ConditionKeys.NotesFoundCount] = (short)count;
        }

        public string AskForSearchTerm()
        {
            SearchTerm = GetDictation();

            return SearchTerm;
        }

        private string GetDictation()
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

        public void PlayBeep()
        {
            SystemSounds.Beep.Play();
        }

        public void PrepareInvoke(string context, Dictionary<string, object> state,
                                  Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            _conditions = conditions;
            _state = state;
            _textValues = textValues;

            if (conditions.ContainsKey(ConditionKeys.NotesFoundCount))
                conditions.Remove(ConditionKeys.NotesFoundCount);

            if (textValues.ContainsKey(TextKeys.SayText))
                conditions.Remove(TextKeys.SayText);

            if (conditions.ContainsKey(ConditionKeys.ErrorCode))
                conditions.Remove(ConditionKeys.ErrorCode);
        }
    }
}
