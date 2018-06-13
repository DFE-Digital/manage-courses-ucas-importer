using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace GovUk.Education.ManageCourses.ApiClient
{
    public class ManageApi
    {
        public void SendToManageCoursesApi(Payload payload)
        {
            Console.WriteLine("Posting to api...");
            var client = new ManageCoursesApiClient();            
            client.ImportAsync(payload).Wait();
            Console.WriteLine("Done.");
        }
    }
}
