using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace VANotes
{
    public class MainPlugin
    {

        private static readonly Guid PluginId = new Guid("120A82D5-A747-446D-B5E7-6EC175B39B6D");
        private static Dictionary<string, IPluginCommand> _pluginCommands;

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
            _pluginCommands = Assembly.GetExecutingAssembly().DefinedTypes
                .Where(t => t.ImplementedInterfaces.Contains(typeof (IPluginCommand)))
                .Select(Activator.CreateInstance)
                .Cast<IPluginCommand>()
                .ToDictionary(i => i.CommandName,
                              new DelegateEqualityComparer<string>((k1,k2) => string.Compare(k1,k2, StringComparison.OrdinalIgnoreCase) == 0,
                                                                   k=>k.GetHashCode()));

            foreach (var command in _pluginCommands)
            {
                command.Value.Init(state, conditions, textValues);
            }
        }

        public static void VA_Exit1(ref Dictionary<string, object> state)
        {
            foreach (var command in _pluginCommands)
            {
                command.Value.Terminate(state);
            }
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

            if (_pluginCommands.ContainsKey(context))
            {
                _pluginCommands[context].Invoke(state, conditions, textValues);
            }
        }


        private static void Clear(Dictionary<string, short?> conditions, Dictionary<string, string> textValues)
        {
            if (conditions.ContainsKey(Keys.NotesFoundKey))
                conditions.Remove(Keys.NotesFoundKey);

            if (textValues.ContainsKey(Keys.NoteResultKey))
                conditions.Remove(Keys.NoteResultKey);
        }
    }
}
