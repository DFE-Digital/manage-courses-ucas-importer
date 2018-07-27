using System;
using System.Threading.Tasks;
using GovUk.Education.ManageCourses.ApiClient;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    public class ManageApi 
    {
        private readonly string apiLocation;
        private readonly string bearerToken;

        public ManageApi(string apiLocation, string bearerToken)
        {
            this.apiLocation = apiLocation;
            this.bearerToken = bearerToken;
        }
        public void PostPayload(UcasPayload payload)
        {
            Console.WriteLine("Posting to api...");
            var client = new ManageCoursesApiClient(new ApiConf(bearerToken), new System.Net.Http.HttpClient());  
            client.BaseUrl = apiLocation;          
            client.Data_ImportAsync(payload).Wait();
            Console.WriteLine("Done.");
        }

        private class ApiConf : IManageCoursesApiClientConfiguration
        {
            private string bearerToken;

            public ApiConf(string bearerToken)
            {
                this.bearerToken = bearerToken;
            }

            public Task<string> GetAccessTokenAsync()
            {
                return Task.FromResult(bearerToken);
            }
        }
    }
}
