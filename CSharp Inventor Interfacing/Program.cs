using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Inventor;

namespace iLogic_Bridge {
    class Program {
        public static Program prog = new Program();
        public static FileSystemWatcher watcher = null;
        public static Dictionary<string, dynamic> nameDocDict = new Dictionary<string, dynamic>();
        public static CommandHandler cmdHandler = new CommandHandler();
        public static OptionsParser options = new OptionsParser();
        public dynamic iLogic;
        public dynamic iLogicAuto;
        public Application invApp;

        Application AttachToInventor() {
            Application app;

            try {
                // Attach to inventor if it's already open
                app = (Application)Marshal.GetActiveObject("Inventor.Application");
            } catch {
                Console.WriteLine("No inventor app found");
                return null;
                // Create an instance of inventor if it hasn't been opened yet
                //Console.WriteLine("Failed to find app, creating instance"); /* DEPRECATED, WE ARE NOT GOING TO START OUR OWN INSTANCE
                //Type appType = Type.GetTypeFromProgID("Inventor.Application");
                //app = (Application)Activator.CreateInstance(appType);
                //app.Visible = false;
                //app.ScreenUpdating = false;
                //isOpen = false;
            }

            return app;
        }

        Object GetiLogicAddIn(Application app) {
            try {
                string iLogicGUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
                ApplicationAddIn iLogicAddIn = app.ApplicationAddIns.ItemById[iLogicGUID];
                iLogicAddIn.Activate();
                return iLogicAddIn;
            } catch (Exception e) {
                Console.WriteLine("Failed to get iLogic Add In");
                Console.WriteLine("Error: {0}", e.Message);
                return null;
            }
        }

        bool SetupInventorConnection() {
            Console.WriteLine("Getting Inventor Application Object...");
            Application ThisApplication = prog.AttachToInventor();
            
            if (ThisApplication != null) {
                prog.invApp = ThisApplication;
                Console.WriteLine("Getting iLogic...");

                prog.iLogic = GetiLogicAddIn(prog.invApp);

                if (prog.iLogic == null) {
                    Console.WriteLine("Failed To Get iLogic AddIn!");
                    return false;
                }

                prog.iLogicAuto = prog.iLogic.Automation;
                prog.iLogicAuto.CallingFromOutside = true;
            } else {
                return false;
            }

            return true;
        }

        static void Main(string[] args) {
            if (prog.SetupInventorConnection()) {
                Console.WriteLine("Setting Up Options...");
                options.OptionsStartup();

                SetupFolder(prog.invApp);
               
                cmdHandler.DisplayHelp();

                try {
                    while (cmdHandler.HandleCommand(prog.invApp, Console.ReadLine())) { }
                } catch (Exception e) {
                    Console.WriteLine("Failure in cmdHandler!");
                    Console.WriteLine("Error: {0}", e.Message);
                    Console.ReadLine();
                }
            }
        }

        public static void SetupFolder(Application ThisApplication) {
            // Surround this all in a try catch, otherwise inventor is entirely locked if it fails
            try {
                prog.invApp.UserInterfaceManager.UserInteractionDisabled = true;

                // Destroy the watcher if it exists
                if (watcher != null) {
                    watcher.Dispose();
                }

                Console.WriteLine("Setting up transfer folder...");

                // Check if our bridge folder exists, and delete it if it does
                if (Directory.Exists(options.options.bridgeFolder)) {
                    Directory.Delete(options.options.bridgeFolder, true);
                }

                // Create the transfer folder
                Directory.CreateDirectory(options.options.bridgeFolder);

                Console.WriteLine("Getting Active Document...");
                Document activeDoc = ThisApplication.ActiveDocument;

                prog.SpanAssemblyTree(options.options.bridgeFolder, activeDoc);

                Console.WriteLine("Creating File Watcher...");
                CreateFileWatcher(options.options.bridgeFolder);

                Console.WriteLine("Watching Files...");

                cmdHandler.isRefreshing = false;
            } catch (COMException e) {
                Console.WriteLine("Something is wrong with the Inventor app object!");
                Console.WriteLine("Error: {0}", e.Message);
                Console.WriteLine("");
                Console.WriteLine("Re-Creating Inventor app object!");
                prog.SetupInventorConnection();
                SetupFolder(prog.invApp);
                prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
            } catch(Exception e) {
                Console.WriteLine("Something went wrong while setting up folder!");
                Console.WriteLine("Error: {0}", e.Message);
                prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
            }
            
            prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
        }

        public void SpanAssemblyTree(string mainPath, Document mainDoc) {
            Console.WriteLine("Creating iLogic File Tree...");

            dynamic rules = prog.iLogicAuto.Rules(mainDoc);
            string assemblyName = mainDoc.DisplayName;
            string curPath = mainPath + "\\" + assemblyName;

            nameDocDict.Clear();
            nameDocDict.Add(assemblyName, mainDoc);

            Directory.CreateDirectory(curPath);
            foreach (dynamic r in rules) {
                System.IO.File.WriteAllText(curPath + "\\" + r.Name + ".vb", r.text);
            }

            // Traverse if it's an assembly and we are going recursive
            if (mainDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject && options.options.recursive) {
                List<string> alreadyFound = new List<string>();
                alreadyFound.Add(mainDoc.DisplayName);
                IterateBranches(curPath, ((AssemblyDocument)mainDoc).ComponentDefinition.Occurrences, ref alreadyFound);
            }
        }

        public void IterateBranches(string curPath, dynamic children, ref List<string> alreadyFound) {
            foreach(ComponentOccurrence child in children) {
                if (child.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject) {
                    try {
                        dynamic doc = child.Definition.Document;

                        // Skip this one if it's been found before
                        if (alreadyFound.Contains(doc.DisplayName)) {
                            continue;
                        }

                        nameDocDict.Add(doc.DisplayName, doc);
                        alreadyFound.Add(doc.DisplayName);
                        dynamic rules = prog.iLogicAuto.Rules(child.Definition.Document);
                        string assemblyName = doc.DisplayName;
                        string newPath = curPath + "\\" + assemblyName;

                        Directory.CreateDirectory(newPath);
                        foreach(dynamic r in rules) {
                            System.IO.File.WriteAllText(newPath + "\\" + r.name + ".vb", r.text);
                        }

                        IterateBranches(newPath, child.SubOccurrences, ref alreadyFound);
                    } catch (Exception) {
                        Console.WriteLine("Failed to enumerate {0}!", child.Name);
                    }
                }
            }
        }

        public static void CreateFileWatcher(string path) {
            // Create a new FileSystemWatcher and set its properties
            watcher = new FileSystemWatcher();
            watcher.Path = path;

            // Watch for changes in LastAccess and LastWrite times, and the renaming of files or directories
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName;
            watcher.Filter = "*.vb";

            // Set it to search sub folders
            watcher.IncludeSubdirectories = true;

            // Add event handlers
            watcher.Changed += OnChanged;
            watcher.Created += OnCreated;
            watcher.Deleted += OnDeleted;
            watcher.Renamed += OnRenamed;

            // Begin watching
            watcher.EnableRaisingEvents = true;
        }

        private static void OnChanged(object source, FileSystemEventArgs e) {
            try {
                // Don't make any changes if we're in the middle of refreshing
                if (!cmdHandler.isRefreshing) {
                    string[] splits = e.Name.Split('\\');
                    string assemblyName = splits[splits.Length - 2];
                    dynamic doc = nameDocDict[assemblyName];

                    string ruleName = splits[splits.Length - 1].Split('.')[0];
                    dynamic rule = prog.iLogicAuto.GetRule(doc, ruleName);

                    if (rule != null) {
                        Console.WriteLine("Updating Rule {0}...", ruleName);
                        string newText = System.IO.File.ReadAllText(e.FullPath);
                        rule.text = newText;
                        RefreshiLogic(doc);
                    } else {
                        Console.WriteLine("**FAILED TO WRITE TO {0} ON CHANGE: RULE DOESNT EXIST**", ruleName);
                    }
                }
            } catch (Exception err) {
                Console.WriteLine("Fatal error somewhere in OnChanged!");
                Console.WriteLine("Error: {0}", err.Message);
            }
        }

        private static void OnCreated(object source, FileSystemEventArgs e) {
            try {
                // Don't make any changes if we're in the middle of refreshing
                if (!cmdHandler.isRefreshing) {
                    string[] splits = e.Name.Split('\\');
                    string assemblyName = splits[splits.Length - 2];
                    dynamic doc = nameDocDict[assemblyName];

                    string ruleName = splits[splits.Length - 1].Split('.')[0];
                    dynamic rule = prog.iLogicAuto.GetRule(doc, ruleName);
                    if (rule == null) {
                        Console.WriteLine("Creating Rule {0}...", ruleName);
                        string newText = System.IO.File.ReadAllText(e.FullPath);
                        rule = prog.iLogicAuto.AddRule(doc, ruleName, "");
                        rule.AutomaticOnParamChange = false;
                        rule.text = newText;
                    } else {
                        Console.WriteLine("**RULE ALREADY EXISTS**", ruleName);
                    }
                }
            } catch (Exception err) {
                Console.WriteLine("Fatal error somewhere in OnCreated!");
                Console.WriteLine("Error: {0}", err.Message);
            }
        }

        private static void OnDeleted(object source, FileSystemEventArgs e) {
            try {
                // Don't make any changes if we're in the middle of refreshing
                if (!cmdHandler.isRefreshing) {
                    string[] splits = e.Name.Split('\\');
                    string assemblyName = splits[splits.Length - 2];
                    dynamic doc = nameDocDict[assemblyName];

                    string ruleName = splits[splits.Length - 1].Split('.')[0];
                    dynamic rule = prog.iLogicAuto.GetRule(doc, ruleName);
                    if (rule != null) {
                        Console.WriteLine("WARNING: Removing Rule {0}...", ruleName);
                        prog.iLogicAuto.DeleteRule(doc, ruleName);
                    }
                }
            } catch (Exception err) {
                Console.WriteLine("Fatal error somewhere in OnDeleted!");
                Console.WriteLine("Error: {0}", err.Message);
            }
        }

        private static void OnRenamed(object source, RenamedEventArgs e) {
            try {
                // Don't make any changes if we're in the middle of refreshing
                if (!cmdHandler.isRefreshing) {
                    if (e.OldName + "~" != e.Name) {
                        // Get the rule name strings
                        string[] oldSplits = e.OldName.Split('\\');
                        string[] newSplits = e.Name.Split('\\');

                        string assemblyName = oldSplits[oldSplits.Length - 2];
                        string oldName = oldSplits[oldSplits.Length - 1].Split('.')[0];
                        string newName = newSplits[newSplits.Length - 1].Split('.')[0];

                        dynamic doc = nameDocDict[assemblyName];

                        // Check and make sure the old rule actually exists
                        dynamic rule = prog.iLogicAuto.GetRule(doc, oldName);

                        if (rule != null) {
                            Console.WriteLine("Renaming Rule {0} To {1}...", oldName, newName);
                            string ruleText = rule.text;
                            prog.iLogicAuto.DeleteRule(doc, oldName);
                            dynamic newRule = prog.iLogicAuto.AddRule(doc, newName, "");
                            newRule.AutomaticOnParamChange = false;
                            newRule.text = ruleText;
                        } else {
                            Console.WriteLine("**OLD RULE {0} DOES NOT EXIST**", oldName);
                        }
                    } else {
                        Console.WriteLine("**SWAP FILE, NOT A RENAME**");
                    }
                }
            } catch (Exception err) {
                Console.WriteLine("Fatal error somewhere in OnRefreshed!");
                Console.WriteLine("Error: {0}", err.Message);
            }
        }

        private static void RefreshiLogic(dynamic doc) {
            prog.iLogicAuto.AddRule(doc, "~!!!~", "");
            prog.iLogicAuto.DeleteRule(doc, "~!!!~");
        }
    }
}
