using System;
using VANotes.Notebooks;

namespace VANotes
{
    public interface IPluginCommand
    {
        /// <summary>
        ///     The name of the command. Passed in via the context from VA to identify the command in the plug-in to execute.
        ///     This is not case-sensitive.
        /// </summary>
        string CommandName { get; }

        /// <summary>
        ///     Initialize any state required by the command plugin.
        /// </summary>
        /// <param name="voiceAttack"></param>
        /// <param name="notebook"></param>
        void Init(VoiceAttack voiceAttack, INotebook notebook);

        /// <summary>
        ///     Terminate the command plug-in and clean up.
        /// </summary>
        void Terminate();

        /// <summary>
        ///     This function is where you will do all of your work. When VoiceAttack encounters an 'Execute External Plug-in Function'
        ///     action, the plug-in indicated will be called.
        /// </summary>
        void Invoke();
    }
}