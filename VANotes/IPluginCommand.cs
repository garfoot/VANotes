using System;
using System.Collections.Generic;

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
        /// <param name="state">
        ///     All values from the state maintained by VoiceAttack for this plug-in. The state allows you to maintain
        ///     kind of a, 'session' within VoiceAttack. This value is not persisted to disk and will be erased on restart.
        ///     Other plug ins do not have access to this state (private to the plug-in).
        /// 
        ///     The state dictionary is the complete state. You can manipulate it however you want,
        ///     the whole thing will be copied back and replace what VoiceAttack is holding on to.
        /// </param>
        /// <param name="conditions">The conditions that were specified in the 'Execute External' command action.</param>
        /// <param name="textValues">The text values that were specified in the 'Execute External' command action.</param>
        void Init(Dictionary<string, object> state, Dictionary<string, Int16?> conditions, Dictionary<string, string> textValues);

        /// <summary>
        ///     Terminate the command plug-in and clean up.
        /// </summary>
        /// <param name="state">
        ///     All values from the state maintained by VoiceAttack for this plug-in. The state allows you to maintain
        ///     kind of a, 'session' within VoiceAttack. This value is not persisted to disk and will be erased on restart.
        ///     Other plug ins do not have access to this state (private to the plug-in).
        /// 
        ///     The state dictionary is the complete state. You can manipulate it however you want,
        ///     the whole thing will be copied back and replace what VoiceAttack is holding on to.
        /// </param>
        void Terminate(Dictionary<string, object> state);

        /// <summary>
        ///     This function is where you will do all of your work. When VoiceAttack encounters an 'Execute External Plug-in Function'
        ///     action, the plug-in indicated will be called.
        /// </summary>
        /// <param name="state">
        ///     All values from the state maintained by VoiceAttack for this plug-in. The state allows you to maintain
        ///     kind of a, 'session' within VoiceAttack. This value is not persisted to disk and will be erased on restart.
        ///     Other plug ins do not have access to this state (private to the plug-in).
        /// 
        ///     The state dictionary is the complete state. You can manipulate it however you want,
        ///     the whole thing will be copied back and replace what VoiceAttack is holding on to.
        /// </param>
        /// <param name="conditions">The conditions that were specified in the 'Execute External' command action.</param>
        /// <param name="textValues">The text values that were specified in the 'Execute External' command action.</param>
        void Invoke(Dictionary<string, object> state, Dictionary<string, Int16?> conditions, Dictionary<string, string> textValues);
    }
}