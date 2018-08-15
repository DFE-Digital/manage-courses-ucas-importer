using System;
using System.Net.Http;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    internal class UcasZipDownloader
    {
        private readonly HttpClient _client;

        public UcasZipDownloader()
        {
            _client = new HttpClient();
        }

        internal void DownloadLatestToFolder(string folder)
        {
        }
    }
}