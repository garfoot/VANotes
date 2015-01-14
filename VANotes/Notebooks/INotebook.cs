using System.Collections.Generic;

namespace VANotes.Notebooks
{
    public interface INotebook
    {
        void ReadNote(VoiceAttack voiceAttack, string pageId);
        IList<Choice> FindPages(string searchTerm);
        string GetPage(string pageId);
    }
}