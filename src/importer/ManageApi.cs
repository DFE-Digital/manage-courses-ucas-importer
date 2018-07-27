using System;
using System.Threading.Tasks;
using GovUk.Education.ManageCourses.ApiClient;

namespace GovUk.Education.ManageCourses.UcasCourseImporter
{
    public class ManageApi 
    {
        private readonly ManageCoursesApiClient client;

        public ManageApi(string apiLocation, string bearerToken)
        {
            client = new ManageCoursesApiClient(new ApiConf(bearerToken), new System.Net.Http.HttpClient());  
            client.BaseUrl = apiLocation;
        }
        
        public void PostPayload(UcasPayload payload)
        {
            Console.WriteLine("Posting to api...");
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
