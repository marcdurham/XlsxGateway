using System;
using System.IO;
using System.Xml;
using XlsxGateway.SystemAdapters;
using XlsxGateway.Models;
using System.Collections.Generic;
using System.Linq;

namespace XlsxGateway.Gateways
{
    public class ExcelXmlWorksheetGateway : IWorksheetGateway
    {
        private const string SharedStringsPath = @"xl/sharedStrings.xml";
        private const string WorkbookPath = @"xl/workbook.xml";
        private const string SheetByIdPath = @"xl/worksheets/sheet{0}.xml";

        private const string GatewayErrorMessage = @"ExcelXmlWorksheetGateway Error: ";
        private const string ErrorSavingSheetMessage = "Error saving Excel XML Worksheet: ";
        private const string ErrorOpeningSheetMessage = "Error opening Excel XML Worksheet: ";
        private const string FileNotFoundMessage = @"File Not Found! File Name: {0}";
        private const string CannotFindZipEntryInFile = @"Cannot find zip archive entry: {0} in file {1}";
        private const string ErrorOpeningZip = @"Error opening Zip File: {0} Error: {1}";
        private const string OpenTargetDocumentErrorMessage = @"Error opening target XML document from zip archive";
        private const string OpenSharedDocumentErrorMessage = @"Error opening a shared XML document from zip archive";
        private const string SaveSharedStringsErrorMessage = @"Error saving shared strings to zip archive: ";
        private const string SaveSheetDocumentErrorMessage = @"Error saving sheet document to zip archive: ";

        Worksheet targetSheet;
        XmlDocument targetSheetDocument;
        Dictionary<string, XmlDocument> otherSheetDocuments = new Dictionary<string, XmlDocument>();

        string targetFileName;
        string targetSheetName;
        ICompressor compressor;
        ISharedStringGateway sharedStringGateway;
        ISheetNameIdGateway sheetNameIdGateway;
        ISheetDocumentSaver sheetDocumentSaver;
        Dictionary<string, int> otherSheetNames;

        public ExcelXmlWorksheetGateway (
            ICompressor compressor,
            ISharedStringGateway sharedStringGateway,
            ISheetNameIdGateway sheetNameIdGateway,
            ISheetDocumentSaver sheetDocumentSaver)
        {
            this.compressor = compressor 
                ?? throw new ArgumentNullException(nameof(compressor));

            this.sharedStringGateway = sharedStringGateway 
                ?? throw new ArgumentNullException(nameof(sharedStringGateway));

            this.sheetNameIdGateway = sheetNameIdGateway 
                ?? throw new ArgumentNullException(nameof(sheetNameIdGateway));

            this.sheetDocumentSaver = sheetDocumentSaver 
                ?? throw new ArgumentNullException(nameof(sheetDocumentSaver));
        }

        // TODO: Open a worksheet from here
        public Worksheet OpenFrom (string fileName, string sheetName)
        {
            return Open (fileName, sheetName);
        }

        public void SaveTo (Worksheet sheet)
        {
            targetFileName = sheet.FileName;
            targetSheetName = sheet.SheetName;

            try
            {
                SaveToArchiveFrom(sheet);
            }
            catch (Exception e) {
                throw new ExcelSheetException (ErrorSavingSheetMessage + e.Message);
            }
        }

        public List<string> SheetNamesFrom(string fileName)
        {
            Open (fileName);
            return otherSheetNames.Keys.ToList();
        }

        private void SaveToArchiveFrom(Worksheet sheet)
        {
            OpenArchiveFrom(sheet.FileName);
            OpenSharedDocuments();
            OpenTargetSheetDocument();
            OpenOtherSheetDocuments();

            // TODO: Does this need to get a return value?
            targetSheetDocument = sheetDocumentSaver.Update(targetSheetDocument, from: sheet);

            sheetDocumentSaver.ResetStringTypes(targetSheetDocument, sheet.StringType);

            ResetStringTypes(otherSheetDocuments.Select(d=>d.Value), toStringType: sheet.StringType);

            SaveSharedStrings();

            SaveSheetDocument(targetSheetDocument, sheet.SheetName, sheet.FileName);
            foreach (var otherDoc in otherSheetDocuments)
                SaveSheetDocument(otherDoc.Value, otherDoc.Key, sheet.FileName);

            compressor.Save();
        }

        private void ResetStringTypes(IEnumerable<XmlDocument> documents, WorksheetStringType toStringType)
        {
            foreach (var document in documents)
                sheetDocumentSaver.ResetStringTypes(document, toStringType);
        }

        void SaveSharedStrings ()
        {
            try {
                var sheetStream = new MemoryStream();

                sharedStringGateway.Save(sheetStream);
                string xml = StringFrom(sheetStream);

                CompressedEntry entry = compressor.GetEntry(SharedStringsPath);
                entry.Content = xml;
            } catch (Exception e) {
                throw new ExcelSheetException (SaveSharedStringsErrorMessage + e.Message);
            }
        }

        void SaveSheetDocument (XmlDocument document, string name, string fileName)
        {
            try {
                var sheetStream = new MemoryStream();

                document.Save (sheetStream);
                string xml = StringFrom(sheetStream);

                CompressedEntry entry = compressor.GetEntry(SheetPathWithIdOf(name, fileName));
                entry.Content = xml;
            } catch (Exception e) {
                throw new ExcelSheetException (SaveSheetDocumentErrorMessage + e.Message);
            }
        }

        static string StringFrom(Stream stream)
        {
            stream.Position = 0;

            string content = new StreamReader(stream)
                .ReadToEnd();

            stream.Close();

            return content;
        }

        string TargetSheetPathWithId ()
        {
            return SheetPathWithIdOf(targetSheetName, targetFileName);
        }

        string SheetPathWithIdOf(string sheetName, string fileName)
        {
            int sheetId = sheetNameIdGateway.SheetIdFrom(sheetName, fileName);

            string sheetPathWithId = string.Format(SheetByIdPath, sheetId);
            return sheetPathWithId;
        }

        void Open (string fileName)
        {
            this.targetFileName = fileName;

            try {
                string filePath = Path.Combine (Environment.CurrentDirectory, fileName);
                compressor.OpenToRead (filePath);
                OpenSharedDocuments ();
                OpenOtherSheetDocuments ();

            } catch (Exception e) {
                throw new ExcelSheetException (ErrorOpeningSheetMessage + e.Message);
            }
        }

        Worksheet Open (string fileName, string sheetName)
        {
            this.targetFileName = fileName;
            this.targetSheetName = sheetName;

            try {
                string filePath = Path.Combine (Environment.CurrentDirectory, fileName);
                compressor.OpenToRead (filePath);
                OpenSharedDocuments ();
                OpenTargetSheetDocument ();
                OpenOtherSheetDocuments ();

            } catch (Exception e) {
                throw new ExcelSheetException (ErrorOpeningSheetMessage + e.Message);
            }

            targetSheet.FileName = this.targetFileName;
            targetSheet.SheetName = this.targetSheetName;

            return targetSheet;
        }

        void OpenSharedDocuments ()
        {
            try {
                sheetNameIdGateway.OpenFrom (OpenXmlFrom (path: WorkbookPath));
                sharedStringGateway.OpenFrom (OpenXmlFrom (path: SharedStringsPath));
            } catch (Exception e) {
                throw new ExcelSheetException (OpenSharedDocumentErrorMessage + e.Message);
            }
        }

        void OpenTargetSheetDocument()
        {
            try
            {
                targetSheetDocument = XmlDocumentFrom(OpenXmlFrom(path: TargetSheetPathWithId()));
                targetSheet = sheetDocumentSaver.OpenFrom(targetSheetDocument); // TODO: Try using the reader here instead of the Saver
                targetSheet.StringType = WorksheetStringType.SharedString;
            }
            catch (Exception e)
            {
                throw new ExcelSheetException(OpenTargetDocumentErrorMessage + e.Message);
            }
        }

        void OpenOtherSheetDocuments()
        {
            try
            {
                otherSheetDocuments.Clear();
                otherSheetNames = sheetNameIdGateway.ExtractSheetNames();
                    
                foreach (var sheetName in otherSheetNames)
                {
                    if (!string.Equals(sheetName.Key, targetSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        var document = XmlDocumentFrom(
                            OpenXmlFrom(path: SheetPathWithIdOf(sheetName.Key, targetFileName)));

                        otherSheetDocuments.Add(sheetName.Key, document);
                    }
                }
            }
            catch (Exception e)
            {
                throw new ExcelSheetException(OpenTargetDocumentErrorMessage + e.Message);
            }
        }

        void OpenArchiveFrom (string fileName)
        {
            if (!File.Exists (fileName))
                throw new System.IO.FileNotFoundException (string.Format(FileNotFoundMessage, fileName));

            try {
                compressor.OpenForUpdate(fileName);
            } catch (Exception e) {
                throw new ExcelSheetException(string.Format(ErrorOpeningZip, fileName, e.Message));
            }
        }

        XmlDocument XmlDocumentFrom(string xml)
        {
            var document = new XmlDocument();
            document.LoadXml(xml);
            return document;
        }

        string OpenXmlFrom (string path)
        {
            var uri = new Uri (path, UriKind.Relative);
            CompressedEntry entry = compressor.GetEntry (uri.ToString ());

            if (entry == null)
                throw new ExcelSheetException (
                    string.Format(CannotFindZipEntryInFile, uri, targetFileName));

            return entry.Content;
        }
    }
}
