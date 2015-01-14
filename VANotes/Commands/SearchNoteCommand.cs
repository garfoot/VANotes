using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VANotes.Notebooks;

namespace VANotes.Commands
{
    public class SearchNoteCommand : IPluginCommand
    {
        private VoiceAttack _voiceAttack;
        private INotebook _notebook;
        
        
        public string CommandName { get { return "SearchNote"; } }
        
        public void Init(VoiceAttack voiceAttack, INotebook notebook)
        {
            _voiceAttack = voiceAttack;
            _notebook = notebook;
        }

        public void Terminate(Dictionary<string, object> state)
        {
        }

        public void Invoke()
        {
            string searchTerm = _voiceAttack.AskForSearchTerm();


            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                _voiceAttack.Error(-1);

                return;
            }

            IList<Choice> choices = _notebook.FindPages(searchTerm);

            _voiceAttack.FoundPages(choices.Count);

            if (choices.Count == 1)
            {
                ReadNote(choices.First().Id);
            }
            else if (choices.Count > 1)
            {
                ReadChoices(choices);
            }
        }

        private void ReadChoices(IList<Choice> choices)
        {
            // Store the choices in VA for later use by other commands
            _voiceAttack.Choices = choices;

            var sb = new StringBuilder();
            sb.AppendFormat("Found {0} pages\r\n", choices.Count);
            
            int i = 1;
            foreach (var choice in choices)
            {
                sb.AppendFormat("Item {0}\r\n {1}\r\n", i++, choice.Name);
            }

            _voiceAttack.Say(sb.ToString());
        }


        private void ReadNote(string pageId)
        {
            _voiceAttack.Say(_notebook.GetPage(pageId));
        }
    }
}