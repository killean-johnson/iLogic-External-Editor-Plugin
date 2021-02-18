using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Inventor;

/*
 * IDEA FOR ILOGIC TREE
 * Start at the top assembly, and get the rules in it.
 * -> These will be at the top level
 * Check the sub assemblies to see if they have any iLogic rules in them
 * -> If they do, put them in a folder under the name of that assembly
 * When a file is modified in any way, check which folder that file is in
 * -> Then get the assembly for that folder
 * --> Then modify the rule appropriately based on that
 * 
 * Changes needed:
 * - Funct that can iterate through each component only ONCE, then throw 
 *   their rules into appropriate folders
 * - The event handlers need to be updated to reflect the folder scheme
 *   and select the proper assembly to update when changes are made
 */

namespace iLogic_Bridge {
    class Program {
        bool isOpen = false;
        static Program prog = new Program();
        static Dictionary<string, dynamic> nameDocDict = new Dictionary<string, dynamic>();
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
                //dynamic rules = prog.iLogicAuto.Rules(activeDoc);
                //foreach (dynamic r in rules) {
                //    System.IO.File.WriteAllText(iLogicTransferFolder + "\\" + r.Name + ".vb", r.text);
                //}
                prog.SpanAssemblyTree(iLogicTransferFolder, activeDoc);

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

        public void SpanAssemblyTree(string mainPath, Document mainDoc) {
            Console.WriteLine("Creating iLogic File Tree...");

            dynamic rules = prog.iLogicAuto.Rules(mainDoc);
            string assemblyName = mainDoc.DisplayName;
            string curPath = mainPath + "\\" + assemblyName;

            nameDocDict.Add(assemblyName, mainDoc);

            Directory.CreateDirectory(curPath);
            foreach (dynamic r in rules) {
                System.IO.File.WriteAllText(curPath + "\\" + r.Name + ".vb", r.text);
            }

            // Traverse if it's an assembly
            if (mainDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) {
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

                        nameDocDict.Add(child._DisplayName, doc);
                        alreadyFound.Add(doc.DisplayName);
                        dynamic rules = prog.iLogicAuto.Rules(child.Definition.Document);
                        string assemblyName = child._DisplayName;
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
            FileSystemWatcher watcher = new FileSystemWatcher();
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
            Console.WriteLine("OnChanged: {0} -|- {1}", e.Name, e.ChangeType);
            //string ruleName = e.Name.Split('.')[0];
            //dynamic rule = prog.iLogicAuto.GetRule(prog.activeDoc, ruleName);

            //if (rule != null) {
            //    Console.WriteLine("Updating Rule {0}...", ruleName);
            //    string newText = System.IO.File.ReadAllText(e.FullPath);
            //    rule.text = newText;
            //} else {
            //    Console.WriteLine("**FAILED TO WRITE TO {0} ON CHANGE: RULE DOESNT EXIST**", ruleName);
            //}
        }

        private static void OnCreated(object source, FileSystemEventArgs e) {
            Console.WriteLine("OnCreated: {0} -|- {1}", e.Name, e.ChangeType);
            // Check to make sure the file name doesn't already exist
            // If it does, skip this rule
            // Otherwise, create the new rule in the document
            //string ruleName = e.Name.Split('.')[0];
            //dynamic rule = prog.iLogicAuto.GetRule(prog.activeDoc, ruleName);

            //if (rule == null) {
            //    Console.WriteLine("Creating Rule {0}...", ruleName);
            //    string newText = System.IO.File.ReadAllText(e.FullPath);
            //    rule = prog.iLogicAuto.AddRule(prog.activeDoc, ruleName, "");
            //    rule.text = newText;
            //} else {
            //    Console.WriteLine("**RULE ALREADY EXISTS**", ruleName);
            //}
        }

        private static void OnDeleted(object source, FileSystemEventArgs e) {
            Console.WriteLine("OnDeleted: {0} -|- {1}", e.Name, e.ChangeType);
            //Console.WriteLine("DeletedFile: {0} {1}", e.Name, e.ChangeType);
            //// Not going to add this functionality unless it's actually needed
            //string ruleName = e.Name.Split('.')[0];
            //dynamic rule = prog.iLogicAuto.GetRule(prog.activeDoc, ruleName);
            //if (rule != null) {
            //    Console.WriteLine("WARNING: Removing Rule {0}...", ruleName);
            //    prog.iLogicAuto.DeleteRule(prog.activeDoc, ruleName);
            //}
        }

        private static void OnRenamed(object source, RenamedEventArgs e) {
            Console.WriteLine("OnRenamed: {0} TO {1}", e.OldFullPath, e.FullPath);

            // SKIP WHEN IT ENDS WITH "~"
            // Specify what is done when a file is renamed.
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
                    newRule.text = ruleText;
                } else {
                    Console.WriteLine("**OLD RULE {0} DOES NOT EXIST**", oldName);
                }
            } else {
                Console.WriteLine("**VIM SWAP FILE, NOT A RENAME**");
            }
        }
    }
}
