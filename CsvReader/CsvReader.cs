using System;
using System.Collections.Generic;
using System.DrawingCore.Text;
using System.IO;
using System.Linq;
using GovUk.Education.ManageCourses.ApiClient;
using GovUk.Education.ManageCourses.Csv.Domain;
using NPOI.SS.Formula.Functions;

namespace GovUk.Education.ManageCourses.Xls
{
    public class CsvReader
    {
        private const int HeaderLine = 1;
        public List<McUser> ReadUsers(string folder)
        {
            var fullFilename = Path.Combine(folder, "mc-users.csv");
            Console.WriteLine("Reading user csv file from: " + fullFilename);

            var users = new List<McUser>();
            var textReader = File.OpenText(fullFilename);
            var csv = new CsvHelper.CsvReader(textReader);
            csv.Read();
            csv.ReadHeader();
            while (csv.Read())
            {
                var record = csv.GetRecord<User>();
                users.Add(new McUser
                {
                    FirstName = record.first_name.Trim(),
                    LastName = record.last_name.Trim(),
                    Email = record.email.Trim()
                });
            }
            Console.Out.WriteLine(users.Count + " users loaded from csv");
            return users;
        }
        public List<McOrganisation> ReadOrganisations(string folder)
        {
            var fullFilename = Path.Combine(folder, "mc-organisations.csv");
            Console.WriteLine("Reading organisation csv file from: " + fullFilename);

            var organisations = new List<McOrganisation>();

            var textReader = File.OpenText(fullFilename);
            var csv = new CsvHelper.CsvReader(textReader);
            csv.Read();
            csv.ReadHeader();
            while (csv.Read())
            {
                var record = csv.GetRecord<Organisation>();
                organisations.Add(new McOrganisation
                {
                    NctlId = record.nctl_id.Trim(),
                    Name = record.name.Trim()
                });
            }
            Console.Out.WriteLine(organisations.Count + " organisations loaded from csv");
            return organisations;
        }
        public List<McOrganisationInstitution> ReadOrganisationInstitutions(string folder)
        {
            var fullFilename = Path.Combine(folder, "mc-organisations_institutions.csv");
            Console.WriteLine("Reading organisation institutions csv file from: " + fullFilename);

            var organisationInstitutions = new List<McOrganisationInstitution>();

            var textReader = File.OpenText(fullFilename);
            var csv = new CsvHelper.CsvReader(textReader);
            csv.Read();
            csv.ReadHeader();
            while (csv.Read())
            {
                var record = csv.GetRecord<OrganisationInstitution>();
                var nctlId = record.nctl_id.Trim();
                var institutionCode = record.institution_code.Trim();
                if (organisationInstitutions.Any(oi => oi.NctlId == nctlId && oi.InstitutionCode == institutionCode))
                {
                    continue; // skip duplicates
                }
                organisationInstitutions.Add(new McOrganisationInstitution
                {
                    NctlId = nctlId,
                    InstitutionCode = institutionCode
                });
            }
   
            Console.Out.WriteLine(organisationInstitutions.Count + " organisation institutions loaded from csv");
            return organisationInstitutions;
        }
        public List<McOrganisationUser> ReadOrganisationUsers(string folder)
        {
            var fullFilename = Path.Combine(folder, "mc-organisations_users.csv");
            Console.WriteLine("Reading organisation users csv file from: " + fullFilename);

            var organisationUsers = new List<McOrganisationUser>();

            var textReader = File.OpenText(fullFilename);
            var csv = new CsvHelper.CsvReader(textReader);
            csv.Read();
            csv.ReadHeader();
            while (csv.Read())
            {
                var record = csv.GetRecord<OrganisationUser>();
                var nctlId = record.nctl_id.Trim();
                var email = record.email.Trim();
                if (organisationUsers.Any(ou => ou.NctlId == nctlId && ou.Email == email))
                {
                    continue; // skip duplicates
                }
                organisationUsers.Add(new McOrganisationUser
                {
                    NctlId = nctlId,
                    Email = email
                });
            }

            Console.Out.WriteLine(organisationUsers.Count + " organisation users loaded from csv");
            return organisationUsers;
        }
    }
}
