using VANotes.Notebooks;

namespace VANotes.Commands
{
    public class GetDictationCommand : IPluginCommand
    {
        private VoiceAttack _voiceAttack;
        public string CommandName { get { return "GetDictation"; } }

        public void Init(VoiceAttack voiceAttack, INotebook notebook)
        {
            _voiceAttack = voiceAttack;
        }

        public void Terminate()
        {
        }

        public void Invoke()
        {
            _voiceAttack.DictationText = _voiceAttack.GetDictation();
        }
    }
}