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
                    string path = Program.options.options.packngoFolder;
                    if (path == "") {
                        path = line.Split('"')[1];
                    }
                    PackNGo(path);
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
