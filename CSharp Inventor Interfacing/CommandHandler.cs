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
                    PackNGo(line.Split('"')[1]);
                    return true;
                case "help":
                    DisplayHelp();
                    return true;
                case "quit":
                    return false;
            }
            return true;
        }

        public void DisplayHelp() {
            Console.WriteLine("Commands:");
            Console.WriteLine("refresh - refresh the folder (This will switch it to whatever project is open)");
            Console.WriteLine("run <rule name> - run the rule in the active document");
            Console.WriteLine("packngo \"Path\" - pack n go the active document, quotes are required");
            Console.WriteLine("help - redisplay the commands");
            Console.WriteLine("quit - end the iLogic bridge");
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
                Program.prog.iLogicAuto.RunRule(doc, ruleName);
            } catch (Exception e) {
                Console.WriteLine("Failed to find rule {0}!", ruleName);
                Console.WriteLine("Error: {0}", e.Message);
            }
        }

        public void PackNGo(string path) {
            string[] fileNameSplits = Program.prog.invApp.ActiveDocument.FullFileName.Split('\\');
            string fileName = fileNameSplits[fileNameSplits.Length - 1];

            Console.WriteLine("Packing {0}...", fileName);
            PackAndGoLib.PackAndGoComponent packNGoComp = new PackAndGoLib.PackAndGoComponent();
            PackAndGoLib.PackAndGo packNGo;

            // Check and see if the directory given even exists
            try {
                if (!Directory.Exists(path)) {
                    Directory.CreateDirectory(path);
                }
            } catch (Exception e) {
                Console.WriteLine("Failed to create directory {0}", path);
                Console.WriteLine("Error: {0}", e.Message);
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
            }

            Console.WriteLine("Setting File Permissions...");
            SetFolderFilePermissions(path);

            try {
                Console.WriteLine("Zipping Folder...");
                ZipFolder(path);
            } catch (Exception e) {
                Console.WriteLine("Failed to zip folder!");
                Console.WriteLine("Error: {0}", e.Message);
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
