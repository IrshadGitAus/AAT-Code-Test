using AAT.File.Processor3.Helpers;
using AAT.File.Processor3.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace AAT.File.Processor3.Services
{
    public class ZipService : IZipService
    {
        private readonly ILogger<ZipService> _logger;
        private readonly IConfiguration _config;
        private ZipConfiguration _zipConfig { get; set; }
        private string _applicationRoot;
        public ZipService(ILogger<ZipService> logger, IConfiguration config, IOptions<ZipConfiguration> zipConfig)
        {
            _logger = logger;
            _config = config;
            _zipConfig = zipConfig.Value;
            _applicationRoot = ApplicationHelper.GetApplicationRoot();
        }

        public bool ProcessFiles()
        {
            bool success = false;
            int count = 0;
            string message = "";
            try
            {
                string inboundFolder = _applicationRoot + "\\" + _zipConfig.InboundFolder;
                string extractFolder = _applicationRoot + "\\" + _zipConfig.ExtractFolder;
                List<string> zipFiles = null;

                if (!Directory.Exists(inboundFolder))
                {
                    _logger.LogError($"Inbound Folder {inboundFolder} does not exist.");
                    throw new Exception();
                }

                if (!Directory.Exists(extractFolder))
                {
                    _logger.LogError($"Extract Folder {extractFolder} does not exist.");
                    throw new Exception();
                }

                zipFiles = GetZipFiles(inboundFolder);
                _logger.LogInformation("Zip files found:" + zipFiles.Count);

                if (zipFiles.Count > 0)
                {
                    zipFiles.ForEach(f =>
                    {
                        count++;
                        _logger.LogInformation($"{Environment.NewLine}");
                        _logger.LogInformation($"{count})Processing the zip file: {f}");
                        if (IsZipFileGood(f))
                        {
                            _logger.LogInformation($"The ZIP file is not corrupt, that is the files inside can be queried.");
                            if (ProcessZipFile(f))
                            {
                                _logger.LogInformation("SUCCESS: ZIP file extracted successfully.");
                            }
                            else
                            {
                                _logger.LogError("FAILURE: ZIP file cannot be extracted.");
                            }
                        }
                        else
                        {
                            _logger.LogError($"FAILURE: ZIP file is corrupt.");
                        }
                    });
                }

                success = true;

                message = $"Congratulations!!.. The zip files are successfully processed. You will find the extracted zip files here {extractFolder}. And see the logs folder for more information.";
                _logger.LogInformation($"{Environment.NewLine}");
                _logger.LogInformation(message);

            }
            catch (Exception)
            {
                success = false;
                return success;
            }

            if (success)
            {
                SendEmailToAdminstrator(message);
            }
            else
            {
                message = "Attention - The zip files were not extracted. Please see the logs folder for more details";
                SendEmailToAdminstrator(message);
            }

            return success;

        }

        public List<string> GetZipFiles(string inboundFolder)
        {
            List<string> zipFiles = Directory.EnumerateFiles(inboundFolder, "*.zip",
                                              SearchOption.TopDirectoryOnly).ToList();

            return zipFiles;
        }

        public static bool IsZipFileGood(string path)
        {
            //The ZIP file is not corrupt, that is the files inside can be queried
            try
            {
                using (var zipFile = ZipFile.OpenRead(path))
                {
                    var entries = zipFile.Entries;
                    return true;
                }
            }
            catch (InvalidDataException)
            {
                return false;
            }
        }

        public bool ProcessZipFile(string zipPath)
        // The ZIP file contains only XLS|X, DOC|X, PDF, MSG and image files(sample documents attached)
        // The ZIP file also contains an XML file called party.XML and it should the structured as per the XSD file attached.
        {
            if (HasAllowedFilesOnly(zipPath))
            {
                _logger.LogInformation("Zip file contains allowed files only.");

                if (ValidateXML(zipPath))
                {
                    _logger.LogInformation("XML file is successfully validated against the XSD file.");

                    if (ExtractZipFile(zipPath))
                    {
                        return true;
                    }
                }
                else
                {
                    _logger.LogError("XML file failed validation against the XSD file.");
                }
            }

            return false;

        }

        public bool HasAllowedFilesOnly(string zipPath)
        {
            try
            {
                using (var zipFile = ZipFile.OpenRead(zipPath))
                {
                    foreach (ZipArchiveEntry e in zipFile.Entries)
                    {
                        if (!IsZipEntryAllowed(e))
                        {
                            _logger.LogError($"{e.FullName} is not a valid entry in the zip file");
                            return false;
                        }
                    }
                }

                return true;
            }
            catch (InvalidDataException)
            {
                return false;
            }
        }

        public bool IsZipEntryAllowed(ZipArchiveEntry zipEntry)
        {
            //The ZIP file contains only XLS| X, DOC | X, PDF, MSG and image files(sample documents attached).
            //The ZIP file also contains an XML file called party.XML and it should the structured as per the XSD file attached.

            string fileName = zipEntry.FullName; //This will hold the full path of an entry inside zip. So, if there are directories then it will hold the path along with '/'
            string extension;

            string[] allowedXMLFiles = _zipConfig.AllowedXMLFiles;
            string[] allowedFileTypes = _zipConfig.AllowedFileTypes;
            string[] allowedImageFileTypes = _zipConfig.AllowedImageFileTypes;

            if (fileName.Contains('/')) return false; //exit if it is a directory

            extension = Path.GetExtension(fileName).TrimStart('.');

            if (!(allowedFileTypes.Contains(extension, StringComparer.OrdinalIgnoreCase) || allowedImageFileTypes.Contains(extension, StringComparer.OrdinalIgnoreCase)))
            {

                if (!allowedXMLFiles.Contains(fileName, StringComparer.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        public bool ValidateXML(string zipPath)
        {
            try
            {
                string xmlFile = _zipConfig.AllowedXMLFiles.Where(x => Path.GetExtension(x).Equals(".xml", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                string xsdFile = _zipConfig.AllowedXMLFiles.Where(x => Path.GetExtension(x).Equals(".xsd", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                Stream xmlFileStream = null;
                Stream xsdFileStream = null;

                using (var zipFile = ZipFile.OpenRead(zipPath))
                {
                    var entries = zipFile.Entries;

                    foreach (ZipArchiveEntry e in zipFile.Entries)
                    {
                        if (e.Name.Equals(xmlFile, StringComparison.OrdinalIgnoreCase))
                        {
                            xmlFileStream = e.Open();
                        }
                        if (e.Name.Equals(xsdFile, StringComparison.OrdinalIgnoreCase))
                        {
                            xsdFileStream = e.Open();
                        }
                    }

                    if (xmlFileStream is null || xsdFileStream is null)
                    {
                        return false;
                    }

                    XmlSchemaSet schema = new XmlSchemaSet();
                    schema.Add("", XmlReader.Create(xsdFileStream));
                    XmlReader rd = XmlReader.Create(xmlFileStream);
                    XDocument doc = XDocument.Load(rd);
                    doc.Validate(schema, null);
                    return true;
                }

            }
            catch (XmlSchemaValidationException)
            {
                return false;
            }
        }

        public bool ExtractZipFile(string zipPath)
        {
            string extractFolder = _zipConfig.ExtractFolder;
            if (string.IsNullOrEmpty(extractFolder))
            {
                _logger.LogError("Extract folder is not valid.");
                return false;
            }
            try
            {
                string zipFolder = GetZipFolder(zipPath);
                extractFolder = _applicationRoot + "\\" + _zipConfig.ExtractFolder + "\\" + zipFolder;
                ZipFile.ExtractToDirectory(zipPath, extractFolder);
                return true;
            }
            catch (IOException e)
            {
                _logger.LogError($"An error occurred while extracting the zip file. ERROR - {e.Message}");
                return false;
            }

        }

        public string GetZipFolder(string zipPath)
        {
            string zipFolder = "";
            string xmlFile = _zipConfig.AllowedXMLFiles.Where(x => Path.GetExtension(x).Equals(".xml", StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            Stream xmlFileStream = null;
            string guid = GetGuid();

            try
            {
                using (var zipFile = ZipFile.OpenRead(zipPath))
                {
                    var entries = zipFile.Entries;

                    var xmlEntry = zipFile.GetEntry(xmlFile);
                    xmlFileStream = xmlEntry.Open();
                    if (xmlFileStream is null)
                    {
                        return Path.GetFileNameWithoutExtension(zipPath) + "-" + guid;
                    }

                    XmlReader rd = XmlReader.Create(xmlFileStream);
                    XDocument doc = XDocument.Load(rd);
                    zipFolder = doc.Elements("party").Single().Element(_zipConfig.XMLPartyAttributeToCreateFolder).Value;
                    zipFolder = zipFolder + "-" + guid;

                    return zipFolder;
                }
            }
            catch (InvalidOperationException)
            {
                return Path.GetFileNameWithoutExtension(zipPath) + "-" + guid;
            }
        }

        public void SendEmailToAdminstrator(string htmlString)
        {
            _logger.LogInformation("Sending mail... please wait.");
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress(_zipConfig.AdministratorEmail);
                message.To.Add(new MailAddress(_zipConfig.AdministratorEmail));
                message.Subject = htmlString;
                message.IsBodyHtml = true; //to make message body as html  
                message.Body = htmlString;
                smtp.Port = 587;
                smtp.Host = "smtp.gmail.com"; //for gmail host  
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("irshad.ahmed@gmail.com", "*******");
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                _logger.LogError($"Email sent to: {_zipConfig.AdministratorEmail}");
            }
            catch (Exception e)
            {
                _logger.LogError($"Sorry.. mail could not be sent. ERROR: {e.Message}");
            }
        }

        private string GetGuid()
        {
            Guid g = Guid.NewGuid();
            return g.ToString();
        }
    }

}
