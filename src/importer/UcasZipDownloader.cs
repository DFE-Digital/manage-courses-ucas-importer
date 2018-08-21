using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Serilog;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    internal class UcasZipDownloader
    {
        private readonly HttpClient _client;
        private readonly ILogger _logger;
        private readonly string _blobContainerUrl;
        private readonly string _sharedAccessSignatureQueryString;

        public UcasZipDownloader(ILogger logger, string blobContainerUrl, string sharedAccessSignature)
        {
            _client = new HttpClient();
            _logger = logger;
            _blobContainerUrl = blobContainerUrl;
            _sharedAccessSignatureQueryString = sharedAccessSignature;
        }

        public async Task<string> DownloadLatestToFolder(string folder)
        {

            // list files
            var listUrl = $"{_blobContainerUrl}?restype=container&comp=list&{_sharedAccessSignatureQueryString}";
            _logger.Debug($"Getting {listUrl}");
            var listResponse = await _client.GetAsync(listUrl);
            var list = XElement.Parse(await listResponse.Content.ReadAsStringAsync());

            var filenames = new List<AzureFile>();
            foreach (var blobElement in list.Element("Blobs").Elements())
            {
                var name = blobElement.Element("Name")?.Value;
                var timestamp = blobElement.Element("Properties")?.Element("Creation-Time")?.Value;
                _logger.Debug($"Parsed blob: Name - '{name}', timestamp - {timestamp}");
                filenames.Add(new AzureFile(name, DateTime.Parse(timestamp)));
            }

            // determine best file
            const string fileNameRegexString = "^NetupdateExtract_[0-9_]+\\.zip$";
            var fileNameRegex = new Regex(fileNameRegexString);
            var matchingFiles = filenames.Where(x => fileNameRegex.IsMatch(x.Name)).ToList();
            _logger.Information($"Found {matchingFiles.Count()} blobs matching '{fileNameRegex}' in azure blob storage");
            var bestFileName = matchingFiles.OrderByDescending(x => x.Timestamp).FirstOrDefault();

            if (bestFileName == null)
            {
                throw new Exception($"Couldn't find any files matching regex '{fileNameRegexString}'");
            }

            // download best file
            string fileToWriteTo = Path.Combine(folder, "ucas-data.zip");

            _logger.Information($"Downloading {bestFileName} to {fileToWriteTo}");

            using (var response = await _client.GetAsync($"{_blobContainerUrl}/{bestFileName.Name}?{_sharedAccessSignatureQueryString}"))
            using (var streamToReadFrom = await response.Content.ReadAsStreamAsync())
            using (Stream streamToWriteTo = File.Open(fileToWriteTo, FileMode.Create))
            {
                await streamToReadFrom.CopyToAsync(streamToWriteTo);
            }
            _logger.Verbose($"Downloading of {bestFileName} to {fileToWriteTo} complete");
            return fileToWriteTo;
        }

        private class AzureFile
        {
            public readonly string Name;
            public readonly DateTime Timestamp;

            public AzureFile(string name, DateTime timestamp)
            {
                Name = name;
                Timestamp = timestamp;
            }

            public override string ToString()
            {
                return $"Azure file {Name} with timestamp {Timestamp}";
            }
        }
    }
}