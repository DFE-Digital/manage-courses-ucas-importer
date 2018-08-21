using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    internal class UcasZipDownloader
    {
        private readonly HttpClient _client;
        private readonly string _blobContainerUrl;
        private readonly string _sharedAccessSignatureQueryString;
        
        public UcasZipDownloader(string blobContainerUrl, string sharedAccessSignature)
        {
            _client = new HttpClient();
            _blobContainerUrl = blobContainerUrl;
            _sharedAccessSignatureQueryString = sharedAccessSignature;
        }

        public async Task<string> DownloadLatestToFolder(string folder)
        {
            string fileToWriteTo = Path.Combine(folder, "ucas-data.zip");
                    
            // list files
            var listResponse = await _client.GetAsync($"{_blobContainerUrl}?restype=container&comp=list&{_sharedAccessSignatureQueryString}");
            var list = XElement.Parse(await listResponse.Content.ReadAsStringAsync());

            var filenames = new List<AzureFile>();
            var fileNameRegex = new Regex("^NetupdateExtract_([0-9]{2})([0-9]{2})([0-9]{4})_([0-9]{2})([0-9]{2})\\.zip$");
            
            foreach(var blobElement in list.Element("Blobs").Elements())
            {
                string name = blobElement.Element("Name").Value;
                var match = fileNameRegex.Match(name);
                if (match.Success) 
                {
                    var timestamp = new DateTime(
                        int.Parse(match.Groups[3].Value),
                        int.Parse(match.Groups[2].Value),
                        int.Parse(match.Groups[1].Value),
                        int.Parse(match.Groups[4].Value),
                        int.Parse(match.Groups[5].Value),
                        0);
                    
                    filenames.Add(new AzureFile(name, timestamp));
                }
            }

            // determine best file
            var bestFileName = filenames.OrderByDescending(x => x.Timestamp).FirstOrDefault();

            if (bestFileName != null) 
            {
                // download best file
                using (HttpResponseMessage response = await _client.GetAsync($"{_blobContainerUrl}/{bestFileName.Name}?{_sharedAccessSignatureQueryString}"))
                using (Stream streamToReadFrom = await response.Content.ReadAsStreamAsync())
                {
                    using (Stream streamToWriteTo = File.Open(fileToWriteTo, FileMode.Create))
                    {
                        await streamToReadFrom.CopyToAsync(streamToWriteTo);
                    }

                    response.Content = null;
                }
                return fileToWriteTo;
            }
            else 
            {
                return null;
            }
        }

        private class AzureFile
        {
            public readonly string Name;
            public readonly DateTime Timestamp;

            public AzureFile(string name, DateTime timestamp)
            {
                this.Name = name;
                this.Timestamp = timestamp;
            }
        }
    }
}