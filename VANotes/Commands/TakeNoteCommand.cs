using VANotes.Notebooks;

namespace VANotes.Commands
{
    public class TakeNoteCommand : IPluginCommand
    {
        private VoiceAttack _voiceAttack;
        private INotebook _notebook;

        public string CommandName { get { return "AddToNote"; } }

        public void Init(VoiceAttack voiceAttack, INotebook notebook)
        {
            _voiceAttack = voiceAttack;
            _notebook = notebook;
        }

        public void Terminate()
        {
        }

        public void Invoke()
        {
            _voiceAttack.NoteText += "\r\n" + _voiceAttack.DictationText;
        }
    }
}