using System;
using System.IO;
using System.Xml;

namespace Downgrade_RDLC
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string folderPath = GetFolderPathFromUser();
            string[] rdlcFiles = GetRDLCFiles(folderPath);

            if (rdlcFiles.Length < 1)
            {
                Console.WriteLine($"No RDLC files found at path {folderPath}");
                ExitApp(1);
            }

            foreach (string rdlc in rdlcFiles)
            {
                if (DowngradeRDLC(rdlc))
                { Console.WriteLine($"Successfully downgraded RDLC: {rdlc}"); }
                else
                { Console.WriteLine($"Failed to downgrade RDLC: {rdlc}"); }
            }

            ExitApp(0);
        }

        private static void ExitApp(int exitCode)
        {
            Console.Write("Press any key to continue . . . ");
            Console.ReadKey();
            Environment.Exit(exitCode);
        }

        private static string GetFolderPathFromUser()
        {
            Console.Write("Input folder path: ");
            string folderPath = Console.ReadLine();

            return folderPath;
        }

        private static string[] GetRDLCFiles(string path)
        {
            try
            {
                var rdlcFiles = Directory.GetFiles(path, "*.rdlc");
                return rdlcFiles;
            }
            catch (Exception)
            { throw; }
        }

        private static bool DowngradeRDLC(string path)
        {
            string xml = GetXml(path);
            var xmlDoc = new XmlDocument();

            xmlDoc.LoadXml(xml);
            XmlElement xmlRoot = xmlDoc.DocumentElement;

            string xmlNamespace = "http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition";

            if (xmlRoot.Attributes["xmlns"].Value != xmlNamespace)
            { xmlRoot.Attributes["xmlns"].Value = xmlNamespace; }

            var nsManager = new XmlNamespaceManager(xmlDoc.NameTable);
            nsManager.AddNamespace("bk", xmlNamespace);

            RemoveChildElements(ref xmlRoot, "AutoRefresh");
            RemoveChildElements(ref xmlRoot, "ReportParametersLayout");

            var reportSections = xmlRoot.GetElementsByTagName("ReportSections");
            if (reportSections.Count > 0)
            {
                var reportSection = reportSections[0].ChildNodes;
                var precedent = reportSections[0];
                foreach (XmlNode child in reportSection[0])
                {
                    var clone = child.Clone();
                    xmlRoot.InsertAfter(clone, precedent);
                    precedent = clone;
                }

                RemoveChildElements(ref xmlRoot, "ReportSections");
            }

            return SaveXmlDocument(ref xmlDoc, path);
        }

        private static string GetXml(string path)
        {
            try
            { return File.ReadAllText(path); }
            catch (Exception)
            { throw; }
        }

        private static bool SaveXmlDocument(ref XmlDocument doc, string path)
        {
            try
            {
                doc.Save(path);
                return true;
            }
            catch (Exception)
            { throw; }
        }

        private static void RemoveChildElements(ref XmlElement documentRoot, string parentElement)
        {
            var elements = documentRoot.GetElementsByTagName(parentElement);
            while (elements.Count > 0)
            {
                documentRoot.RemoveChild(elements[0]);
            }
        }
    }
}
