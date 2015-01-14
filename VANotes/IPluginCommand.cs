using System;
using System.Collections.Generic;

namespace VANotes
{
    public interface IPluginCommand
    {
        string CommandName { get; }

        void Init(Dictionary<string, object> state, Dictionary<string, Int16?> conditions, Dictionary<string, string> textValues);

        void Terminate(Dictionary<string, object> state);

        void Invoke(Dictionary<string, object> state, Dictionary<string, Int16?> conditions, Dictionary<string, string> textValues);
    }
}