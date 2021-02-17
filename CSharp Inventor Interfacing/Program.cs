using System;
using System.IO;
using System.Runtime.InteropServices;
using Inventor;

namespace CSharp_Inventor_Interfacing {
    class Program {
        bool isOpen = false;
        static Program prog = new Program();
        dynamic iLogicAuto;
        dynamic activeDoc;

        Application AttachToInventor() {
            Application app;

            try {
                // Attach to inventor if it's already open
                app = (Application)Marshal.GetActiveObject("Inventor.Application");
                isOpen = true;
            } catch {
                // Create an instance of inventor if it hasn't been opened yet
                Console.WriteLine("Failed to find app, creating instance");
                Type appType = Type.GetTypeFromProgID("Inventor.Application");
                app = (Application)Activator.CreateInstance(appType);
                app.Visible = false;
                app.ScreenUpdating = false;
                isOpen = false;
            }

            if (app == null) {
                Console.WriteLine("No inventor app found, and it was failed to be created");
                return null;
            }

            return app;
        }

        Object GetiLogicAddIn(Application app) {
            string iLogicGUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
            ApplicationAddIn iLogicAddIn = app.ApplicationAddIns.ItemById[iLogicGUID];
            iLogicAddIn.Activate();
            return iLogicAddIn.Automation;
        }

        static void Main(string[] args) {
            // The folder where we'll be doing our work from
            const string iLogicTransferFolder = "C:\\iLogicTransfer";

            // Check if it exists, and delete it if it does
            if (Directory.Exists(iLogicTransferFolder)) {
                Directory.Delete(iLogicTransferFolder, true);
            }

            // Create the transfer folder
            Directory.CreateDirectory(iLogicTransferFolder);

            Console.WriteLine("Getting Inventor Application Object...");
            Application ThisApplication = prog.AttachToInventor();

            if (ThisApplication != null) {
                Console.WriteLine("Getting iLogic...");
                prog.iLogicAuto = prog.GetiLogicAddIn(ThisApplication);
                prog.iLogicAuto.CallingFromOutside = true;

                Console.WriteLine("Getting Active Document...");
                Document activeDoc = ThisApplication.ActiveDocument;
                prog.activeDoc = activeDoc;

                Console.WriteLine("Creating iLogic Files In iLogicTransfer Folder...");
                dynamic rules = prog.iLogicAuto.Rules(activeDoc);
                foreach (dynamic r in rules) {
                    System.IO.File.WriteAllText(iLogicTransferFolder + "\\" + r.Name + ".vb", r.text);
                }

                Console.WriteLine("Creating File Watcher...");
                CreateFileWatcher(iLogicTransferFolder);

                if (!prog.isOpen) {
                    Console.WriteLine("Quitting Inventor...");
                    ThisApplication.Quit();
                }

                Console.WriteLine("Watching Files...");

                Console.ReadLine();
            }
        }
        public static void CreateFileWatcher(string path) {
            // Create a new FileSystemWatcher and set its properties
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = path;
            // Watch for changes in LastAccess and LastWrite times, and the renaming of files or directories
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName;
            watcher.Filter = "*.vb";

            // Add event handlers
            watcher.Changed += OnChanged;
            watcher.Created += OnCreated;
            watcher.Deleted += OnDeleted;
            watcher.Renamed += OnRenamed;

            // Begin watching
            watcher.EnableRaisingEvents = true;
        }

        private static void OnChanged(object source, FileSystemEventArgs e) {
            string ruleName = e.Name.Split('.')[0];
            dynamic rule = prog.iLogicAuto.GetRule(prog.activeDoc, ruleName);

            if (rule != null) {
                Console.WriteLine("Updating Rule {0}...", ruleName);
                string newText = System.IO.File.ReadAllText(e.FullPath);
                rule.text = newText;
            } else {
                Console.WriteLine("**FAILED TO WRITE TO {0} ON CHANGE: RULE DOESNT EXIST**", ruleName);
            }
        }
        private static void OnCreated(object source, FileSystemEventArgs e) {
            // Check to make sure the file name doesn't already exist
            // If it does, skip this rule
            // Otherwise, create the new rule in the document
            string ruleName = e.Name.Split('.')[0];
            dynamic rule = prog.iLogicAuto.GetRule(prog.activeDoc, ruleName);

            if (rule == null) {
                Console.WriteLine("Creating Rule {0}...", ruleName);
                string newText = System.IO.File.ReadAllText(e.FullPath);
                rule = prog.iLogicAuto.AddRule(prog.activeDoc, ruleName, "");
                rule.text = newText;
            } else {
                Console.WriteLine("**RULE ALREADY EXISTS**", ruleName);
            }
        }

        private static void OnDeleted(object source, FileSystemEventArgs e) {
            Console.WriteLine("DeletedFile: {0} {1}", e.Name, e.ChangeType);
            // Not going to add this functionality unless it's actually needed
            string ruleName = e.Name.Split('.')[0];
            dynamic rule = prog.iLogicAuto.GetRule(prog.activeDoc, ruleName);
            if (rule != null) {
                Console.WriteLine("WARNING: Removing Rule {0}...", ruleName);
                prog.iLogicAuto.DeleteRule(prog.activeDoc, ruleName);
            }
        }

        private static void OnRenamed(object source, RenamedEventArgs e) {
            // SKIP WHEN IT ENDS WITH "~"
            // Specify what is done when a file is renamed.
            if (e.OldName + "~" != e.Name) {
                
            } else {
                Console.WriteLine("**VIM SWAP FILE, NOT A RENAME**");
            }
        }
    }
}
