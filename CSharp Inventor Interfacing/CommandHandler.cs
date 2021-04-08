using System;
using System.Collections.Generic;
using System.IO;
using Inventor;

namespace iLogic_Bridge {
    class CommandHandler {
        public bool isRefreshing = false;

        public bool HandleCommand(Application ThisApplication, string line) {
            string[] splits = line.Split();
            switch (splits[0]) {
                case "refresh":
                    isRefreshing = true;
                    Program.SetupFolder(ThisApplication);
                    return true;

                case "run":
                    RunRuleCommand(line);
                    return true;

                case "packngo":
                    string packPath = Program.options.options.packngoFolder;
                    if (packPath == "") {
                        packPath = line.Split('"')[1];
                    }
                    PackNGo(packPath);
                    return true;

                case "store":
                    string storePath = Program.options.options.storageFolder;
                    if (storePath == "") {
                        storePath = line.Split('"')[1];
                    }
                    Store(storePath);
                    return true;

                case "set":
                    Program.options.SetOptions(line);
                    break;

                case "showoptions":
                    Program.options.ShowOptions();
                    break;

                case "help":
                    DisplayHelp();
                    return true;

                case "quit":
                    return false;
            }
            return true;
        }

        public void DisplayHelp() {
            Console.Clear();
            Console.WriteLine("Commands:");
            Console.WriteLine("refresh - refresh the folder (This will switch it to whatever project is open)");
            Console.WriteLine("run <rule name> - run the rule in the active document");
            Console.WriteLine("packngo \"Path\" - pack n go the active document, quotes are required. The path is optional if one \n\tis set already in the options");
            Console.WriteLine("store - store the active document's ilogic rules (currenly open) in the storage folder. These don't update the document when changed, it's just for reference");
            Console.WriteLine("showoptions - show the current value of each option");
            Console.WriteLine("set <option> <value> - sets an option to the specified value and writes it to the options json file. \n\tPaths must be surrounded by quotes\n\tType \"set help\" for a list of options");
            Console.WriteLine("help - redisplay the commands");
            Console.WriteLine("quit - end the iLogic bridge");
        }

        public static void DisplayOptionsHelp() {
            Console.Clear();
            Console.WriteLine("Options:");
            Console.WriteLine("recursive - Can be true or false. If true, iterates through child assemblies and lets you edit their iLogic if it \n\texists. If it's false, it only covers the active document assembly");
            Console.WriteLine("blocking - Can be true or false. If true, the run command will block input to inventor during the course of the rule \n\tbeing ran, which significantly speeds up run time");
            Console.WriteLine("bridgefolder - Sets the path for where iLogic rules are stored for editing. THE PATH MUST BE SURROUNDED BY QUOTES");
            Console.WriteLine("packngofolder - Sets the default path for where the results of the packngo command are stored. \n\tTHE PATH MUST BE SURROUNDED BY QUOTES");
            Console.WriteLine("storagefolder - Sets the default path for where the iLogic rules are stored. \n\tTHE PATH MUST BE SURROUNDED BY QUOTES");
        }

        public void RunRuleCommand(string line) {
            string[] splits = line.Split();
            string ruleName = "";
            for (int i = 1; i < splits.Length; i++) {
                ruleName += splits[i];
                if (i != splits.Length - 1)
                    ruleName += " ";
            }

            dynamic doc = Program.prog.invApp.ActiveDocument;

            try {
                Program.prog.invApp.UserInterfaceManager.UserInteractionDisabled = Program.options.options.blocking;
                Program.prog.iLogicAuto.RunRule(doc, ruleName);
            } catch (Exception e) {
                Console.WriteLine("Failed to find rule {0}!", ruleName);
                Console.WriteLine("Error: {0}", e.Message);
                Program.prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
            }
            Program.prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
        }

        public void PackNGo(string path) {
            string[] fileNameSplits = Program.prog.invApp.ActiveDocument.FullFileName.Split('\\');
            string fileName = fileNameSplits[fileNameSplits.Length - 1];

            Console.WriteLine("Packing {0}...", fileName);
            PackAndGoLib.PackAndGoComponent packNGoComp = new PackAndGoLib.PackAndGoComponent();
            PackAndGoLib.PackAndGo packNGo;

            // Make sure the path is absolute
            if (!System.IO.Path.IsPathRooted(path)) {
                Console.WriteLine("PackNGo Path Needs To Be Absolute");
                return;
            }

            // Check and see if the directory given even exists
            try {
                if (!Directory.Exists(path)) {
                    Directory.CreateDirectory(path);
                }
            } catch (Exception e) {
                Console.WriteLine("Failed to create directory {0}", path);
                Console.WriteLine("Error: {0}", e.Message);
                return;
            }

            try {
                packNGo = packNGoComp.CreatePackAndGo(Program.prog.invApp.ActiveEditDocument.FullFileName, path);
                packNGo.ProjectFile = Program.prog.invApp.DesignProjectManager.ActiveDesignProject.FullFileName;

                packNGo.SkipLibraries = true;
                packNGo.SkipStyles = true;
                packNGo.SkipTemplates = true;
                packNGo.CollectWorkgroups = false;
                packNGo.KeepFolderHierarchy = false;
                packNGo.IncludeLinkedFiles = true;

                string[] refFiles;
                object missFiles;
                packNGo.SearchForReferencedFiles(out refFiles, out missFiles);
                packNGo.AddFilesToPackage(ref refFiles);

                packNGo.CreatePackage();
            } catch (Exception e) {
                Console.WriteLine("Failed to create package!");
                Console.WriteLine("Error: {0}", e.Message);
                return;
            }

            Console.WriteLine("Setting File Permissions...");
            SetFolderFilePermissions(path);

            try {
                Console.WriteLine("Zipping Folder...");
                ZipFolder(path);
            } catch (Exception e) {
                Console.WriteLine("Failed to zip folder!");
                Console.WriteLine("Error: {0}", e.Message);
                return;
            }

            Console.WriteLine("Finished Packing!");
        }

        public void Store(string path) {
            // Surround this in a try/catch in case it fails
            try {
                Program.prog.invApp.UserInterfaceManager.UserInteractionDisabled = true;

                Console.WriteLine("Getting active document...");
                Document activeDoc = Program.prog.invApp.ActiveDocument;

                Console.WriteLine("Setting up storage folder...");
                Program.prog.SpanAssemblyTree(Program.options.options.storageFolder, activeDoc);
                
            } catch (Exception e) {
                Console.WriteLine("Something went wrong while setting up folder!");
                Console.WriteLine("Error: {0}", e.Message);
                Program.prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
            }

            Console.WriteLine("Finished setting up storage!");
            Program.prog.invApp.UserInterfaceManager.UserInteractionDisabled = false;
        }

        private void StorageSpanTree(string mainPath, Document mainDoc) {
            Console.WriteLine("Creating iLogic File Tree...");

            dynamic rules = Program.prog.iLogicAuto.Rules(mainDoc);
            string assemblyName = mainDoc.DisplayName;
            string curPath = mainPath + "\\" + assemblyName;
            
            // Check if our bridge folder exists, and delete it if it does
            if (Directory.Exists(curPath)) {
                Directory.Delete(curPath, true);
            }

            // Create the transfer folder
            Directory.CreateDirectory(curPath);

            Directory.CreateDirectory(curPath);
            foreach (dynamic r in rules) {
                System.IO.File.WriteAllText(curPath + "\\" + r.Name + ".vb", r.text);
            }

            // Traverse if it's an assembly and we are going recursive
            if (mainDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject && Program.options.options.recursive) {
                List<string> alreadyFound = new List<string>();
                alreadyFound.Add(mainDoc.DisplayName);
                StorageIterateBranches(curPath, ((AssemblyDocument)mainDoc).ComponentDefinition.Occurrences, ref alreadyFound);
            }
        }

        private void StorageIterateBranches(string curPath, dynamic children, ref List<string> alreadyFound) {
            foreach(ComponentOccurrence child in children) {
                if (child.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject) {
                    try {
                        dynamic doc = child.Definition.Document;

                        // Skip this one if it's been found before
                        if (alreadyFound.Contains(doc.DisplayName)) {
                            continue;
                        }

                        alreadyFound.Add(doc.DisplayName);
                        dynamic rules = Program.prog.iLogicAuto.Rules(child.Definition.Document);
                        string assemblyName = doc.DisplayName;
                        string newPath = curPath + "\\" + assemblyName;

                        Directory.CreateDirectory(newPath);
                        foreach(dynamic r in rules) {
                            System.IO.File.WriteAllText(newPath + "\\" + r.name + ".vb", r.text);
                        }

                        StorageIterateBranches(newPath, child.SubOccurrences, ref alreadyFound);
                    } catch (Exception) {
                        Console.WriteLine("Failed to enumerate {0}!", child.Name);
                    }
                }
            }
        }

        public void SetFolderFilePermissions(string path) {
            var files = traverse(path);
            foreach (string file in files) {
                FileInfo fileDetail = new FileInfo(file);
                fileDetail.IsReadOnly = false;
            }
        }

        private IEnumerable<string> traverse(string path) {
            foreach (string f in Directory.GetFiles(path)) {
                yield return f;
            }

            foreach (string d in Directory.GetDirectories(path)) {
                foreach (string f in traverse(d)) {
                    yield return f;
                }
            }
        }

        private void ZipFolder(string path) {
            string[] splitPath = path.Split('\\');
            string zipPath = string.Join("\\", splitPath, 0, splitPath.Length - 1) + "\\Generator.zip";

            // Delete the zip file if it already exists
            if (System.IO.File.Exists(zipPath)) {
                System.IO.File.Delete(zipPath);
            }

            // Zip the folder
            System.IO.Compression.ZipFile.CreateFromDirectory(path, zipPath);
        }
    }
}
