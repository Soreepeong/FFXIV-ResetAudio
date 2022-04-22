using Dalamud.Configuration;
using Dalamud.Plugin;
using System;
using System.Collections.Generic;
using System.Linq;
using static ResetAudio.CantBelieveItsNotCpp;

namespace ResetAudio {
    [Serializable]
    public class Configuration : IPluginConfiguration {
        public int Version { get; set; } = 0;

        public bool ConfigVisible = true;

        public bool AdvanceConfigExpanded = false;

        public bool SuppressMultimediaDeviceChangeNotificationToGame = true;

        public bool PrintAudioResetToChat = true;

        public bool Patch156e4e3 = true;

        public int CoalesceAudioResetRequestDurationMs = 100;

        public List<PropertyKeyWithComment> IgnorePropertyUpdateKeys = new();

        // the below exist just to make saving less cumbersome

        [NonSerialized]
        private DalamudPluginInterface? _pluginInterface;

        public void Initialize(DalamudPluginInterface pluginInterface) {
            _pluginInterface = pluginInterface;

            if (!IgnorePropertyUpdateKeys.Any()) {
                IgnorePropertyUpdateKeys.Add(
                    new() {
                        PropKey = new() {
                            FmtId = new(0x9855c4cd, 0xdf8c, 0x449c, 0xa1, 0x81, 0x81, 0x91, 0xb6, 0x8b, 0xd0, 0x6c),
                            Pid = 0,
                        },
                        Comment = "Audio client attach?",
                        Enable = true,
                    }
                );
            }
        }

        public void Save() {
            _pluginInterface!.SavePluginConfig(this);
        }

        [Serializable]
        public class PropertyKeyWithComment {
            public PropertyKey PropKey;
            public string Comment = "";
            public bool Enable = true;
        }
    }
}
