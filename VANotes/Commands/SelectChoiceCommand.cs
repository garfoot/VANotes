using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Speech.Recognition;
using VANotes.Notebooks;

namespace VANotes.Commands
{
    public class SelectChoiceCommand : IPluginCommand
    {
        private INotebook _notebook;
        private VoiceAttack _voiceAttack;
        private const double ConfidenceLevel = 0.0;
        
        
        public string CommandName { get { return "SelectChoice"; }}

        public void Init(VoiceAttack voiceAttack, INotebook notebook)
        {
            _notebook = notebook;
            _voiceAttack = voiceAttack;
        }

        public void Terminate(Dictionary<string, object> state)
        {
        }

        public void Invoke()
        {
            var choices = _voiceAttack.Choices;

            var grammar = new GrammarBuilder();
            string[] choicesIndexes = Enumerable.Range(1, choices.Count).Select(i => i.ToString()).ToArray();

            grammar.Append(new Choices(choicesIndexes));

            using (var recognizer = new SpeechRecognitionEngine(CultureInfo.CurrentUICulture))
            {
                recognizer.LoadGrammar(new Grammar(grammar));
                recognizer.SetInputToDefaultAudioDevice();

                _voiceAttack.PlayBeep();

                var result = recognizer.Recognize();

                if (result != null && result.Confidence > ConfidenceLevel)
                {
                    string res = result.Text;
                    int index;
                    if (int.TryParse(res, out index))
                    {
                        _notebook.ReadNote(_voiceAttack, choices[index - 1].Id);
                    }
                }
            }
        }
    }
}