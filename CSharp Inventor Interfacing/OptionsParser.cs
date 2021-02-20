using System;
using System.Text.Json;

namespace iLogic_Bridge {
    class OptionsParser {
        public struct Options {
            public bool recursive;
            public bool blocking;
            public string bridgeFolder;
            public string packngoFolder;
        }

        private const string optionsFilePath = "iLogicBridgeOptions.json";
        public Options options;

        public void OptionsStartup() {
            // Check to see if the options file exists
            if (System.IO.File.Exists(optionsFilePath)) {
                Console.WriteLine("Parsing Options File...");
                ParseOptions();
            } else {
                Console.WriteLine("Creating New Options File...");
                CreateDefaultsFile();
            }
        }

        void CreateDefaultsFile() {
            options.recursive = false;
            options.blocking = true;
            options.bridgeFolder = "C:\\iLogicBridge";
            options.packngoFolder = "";

            string optionsJson = JsonSerializer.Serialize(options);
            System.IO.File.WriteAllText(optionsFilePath, optionsJson);
        }

        void ParseOptions() {
            string jsonString = System.IO.File.ReadAllText(optionsFilePath);
            options = JsonSerializer.Deserialize<Options>(jsonString);
        }

        void WriteOutOptions() {
            string optionsJson = JsonSerializer.Serialize(options);
            System.IO.File.WriteAllText(optionsFilePath, optionsJson);
        }

    }
}
