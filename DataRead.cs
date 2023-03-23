using Microsoft.Graph;
using OfficeOpenXml;
using System.Diagnostics;

namespace cert_mailer
{
    public class DataRead
    {
        const int GRADESPACE = 1;
        const int ROSTERSPACE = 11;
        private Dictionary<string, string> certMap = new Dictionary<string, string>();
        private string certPath;
        public Course course { get; }

        public DataRead(FileInfo roster, FileInfo grades, String certPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var rosterExcel = new ExcelPackage(roster);
            // Get the Worksheet for the Roster
            using var rosterSheet = rosterExcel.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name.Contains("EOC"));
            if (rosterSheet == null)
            {
                throw new Exception("The worksheet with name containing 'EOC' was not found.");
            }
            // Get the Worksheet for the Grades
            using var gradesExcel = new ExcelPackage(grades);
            using var gradesSheet = gradesExcel.Workbook.Worksheets[0];
            // Set CertPath
            this.certPath = certPath;
            // Set Course
            string fileName = grades.Name;
            course = courseReader(rosterSheet, gradesSheet, fileName);
            studentReader(rosterSheet, gradesSheet);
        }

        public void studentReader(ExcelWorksheet RosterSheet, ExcelWorksheet GradesSheet)
        {
            // Compare the data in each row
            var rowCount = Math.Max(RosterSheet.Dimension.End.Row, GradesSheet.Dimension.End.Row);
            for (var row = 1; row <= rowCount; row++)
            {
                int rosterSpacing = row + ROSTERSPACE;
                var rosterFirstName = RosterSheet.Cells[rosterSpacing, 2].Value?.ToString();
                var rosterLastName = RosterSheet.Cells[rosterSpacing, 3].Value?.ToString();
                var rosterEmail = RosterSheet.Cells[rosterSpacing, 4].Value?.ToString();

                int gradeSpacing = row + GRADESPACE;
                var gradesFirstName = GradesSheet.Cells[gradeSpacing, 3].Value?.ToString();
                var gradesLastName = GradesSheet.Cells[gradeSpacing, 5].Value?.ToString();
                var gradesGrade = GradesSheet.Cells[gradeSpacing, 6].Value?.ToString();

                string firstName = rosterFirstName ?? "";
                string lastName = rosterLastName ?? "";
                string email = rosterEmail ?? "";
                string grade = gradesGrade ?? "";

                string fullName = $"{firstName} {lastName}";
                string cert = MatchCert(fullName);

                if (rosterFirstName == gradesFirstName && rosterLastName == gradesLastName && rosterFirstName != null && GradeCheck(grade))
                {
                    Student student = new Student(firstName, lastName, email, cert, grade);
                    course.AddStudent(student);
                }
            }
        }
        private bool GradeCheck(string grade)
        {
            if (grade.ToLower() == "pass")
            {
                return true;
            }
            double result;
            bool isDouble = double.TryParse(grade, out result);
            if (isDouble && result >= 0.8)
            {
                return true;
            }
            return false;
        }

        public Course courseReader(ExcelWorksheet rosterSheet, ExcelWorksheet gradesSheet, string fileName)
        {
            var courseName = rosterSheet.Cells[5, 2].Value?.ToString();
            courseName = colonSplit(courseName ?? "");
            var courseId = gradesSheet.Cells[2, 1].Value?.ToString();
            var agency = gradesSheet.Cells[2, 10].Value?.ToString();
            var instructor = rosterSheet.Cells[7, 10].Value?.ToString();
            instructor = colonSplit(instructor ?? "");
            var location = gradesSheet.Cells[2, 11].Value?.ToString();
            var startDateString = gradesSheet.Cells[2, 8].Value?.ToString();
            var endDateString = gradesSheet.Cells[2, 9].Value?.ToString();
            DateTime startDate = stringDateTime(startDateString ?? "");
            DateTime endDate = stringDateTime(endDateString ?? "");

            // Get a possible courseId override
            string courseId2 = fileName;
            courseId2 = courseId2.Replace("BMRA Roster and Grades - ", "");
            courseId2 = courseId2.Replace(".xlsx", "");

            if (courseId != courseId2)
            {
                courseId = courseId2;
            }

            return new Course(courseName ?? "", courseId ?? "", agency ?? "", instructor ?? "", location ?? "", startDate, endDate);
        }


        private DateTime stringDateTime(string date)
        {
            return DateTime.Parse(date);
        }

        private string colonSplit(string input)
        {
            string[] parts = input.Split(": ");
            string a = parts[0];
            string b = parts[1];
            return b;
        }

        private void BuildCertMap()
        {
            // Get all the PDF files in the specified directory
            string[] files = System.IO.Directory.GetFiles(certPath, "*.pdf");

            // Loop through each file
            foreach (string file in files)
            {
                // Get the file name without the extension and split it by the delimiter
                string fileName = Path.GetFileNameWithoutExtension(file);
                string[] nameParts = fileName.Split(" - ");

                // Make sure the file name has at least two parts (full name and certificate type)
                if (nameParts.Length >= 2)
                {
                    // Get the full name from the file name and add it to the certMap if it's not already there
                    string fullName = nameParts[1];
                    if (!certMap.ContainsKey(fullName))
                    {
                        certMap.Add(fullName, file);
                    }
                }
            }
        }

        private string MatchCert(string searchString)
        {
            if (certMap.Count == 0)
            {
                BuildCertMap();
            }

            try
            {
                return certMap[searchString];
            }
            catch (KeyNotFoundException)
            {
                Console.WriteLine("Could not find certificate for" + searchString);
                return "Could not find certificate for" + searchString;
            }
        }

    }
}