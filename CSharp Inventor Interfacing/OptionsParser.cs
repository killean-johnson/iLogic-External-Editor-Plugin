﻿using System;
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
        public void WriteOutOptions() {
            string optionsJson = JsonSerializer.Serialize(options);
            System.IO.File.WriteAllText(optionsFilePath, optionsJson);
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

        public void SetOptions(string line) {
            string[] splits = line.Split();

            switch (splits[1]) {
                case "recusive":
                    bool recursive = options.recursive;
                    switch (splits[2].ToLower()) {
                        case "true":
                            recursive = true;
                            break;
                        case "false":
                            recursive = false;
                            break;
                        default:
                            Console.WriteLine("{0} is not a valid setting for recursive. Enter in either true or false");
                            break;
                    }

                    options.recursive = recursive;
                    WriteOutOptions();
                    break;

                case "blocking":
                    bool blocking = options.blocking;
                    switch (splits[2].ToLower()) {
                        case "true":
                            blocking = true;
                            break;
                        case "false":
                            blocking = false;
                            break;
                        default:
                            Console.WriteLine("{0} is not a valid setting for blocking. Enter in either true or false");
                            break;
                    }

                    options.blocking = blocking;
                    WriteOutOptions();
                    break;

                case "bridgefolder":
                    string path = line.Split('"')[1];
                    options.bridgeFolder = path;
                    WriteOutOptions();
                    break;

                case "packngofolder":
                    string path = line.Split('"')[1];
                    options.packngoFolder = path;
                    WriteOutOptions();
                    break;

                case "help":
                    CommandHandler.DisplayOptionsHelp();
                    break;

                default:
                    Console.WriteLine("{0} not a registered option", splits[1]);
                    break;
            }
        }
        
        public void ShowOptions() {
            Console.WriteLine("Option Values: ");
            Console.WriteLine("recursive: {0}", options.recursive);
            Console.WriteLine("blocking: {0}", options.blocking);
            Console.WriteLine("bridgefolder: {0}", options.bridgeFolder);
            Console.WriteLine("packngofolder: {0}", options.packngoFolder);
        }
    }
}