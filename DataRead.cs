using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;

namespace cert_mailer
{
    public class DataRead
    {
        private readonly Dictionary<int, string> rosterType = new()
        {
            {1, "BMRA"},
            {2, "VA"},
            {3, "OTHER"}
        };
        public const int GRADESPACE = 1;
        const int BMRAROSTERSPACE = 11;
        private readonly Dictionary<string, string> certMap = new();
        public string certPath;
        private int minimumPassingGrade;
        public Course Course { get; }

        public DataRead(FileInfo roster, FileInfo grades, string certPath, bool createCerts, EnumCertificateType.CertificateType certificateType, int minimumPassingGrade)
        {
            // Set CertPath
            this.certPath = certPath;

            int rosterType = 1;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var rosterExcel = new ExcelPackage(roster);
            // Get the Worksheet for the Roster
            // If it's a BMRA roster
            ExcelWorksheet? rosterSheet = null;
            if (IsBMRARoster(rosterExcel))
            {
                rosterSheet = rosterExcel.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name.Contains("EOC"));
                rosterType = 1;
                // Ensure that there's an EOC page
                if (rosterSheet == null)
                {
                    throw new Exception("The worksheet with name containing 'EOC' was not found.");
                }
            }
            // If it's not BMRA it's gotta be the VA! idk...
            else
            {
                foreach (ExcelWorksheet worksheet in rosterExcel.Workbook.Worksheets)
                {
                    if (worksheet.Hidden == eWorkSheetHidden.Visible)
                    {
                        rosterSheet = worksheet;
                    }
                }
                rosterType = 2;
            }

            if (rosterSheet == null)
            {
                throw new Exception("The worksheet was not found.");
            }

            // min passing grade
            this.minimumPassingGrade = minimumPassingGrade;

            // Get the Worksheet for the Grades
            using var gradesExcel = new ExcelPackage(grades);
            using var gradesSheet = gradesExcel.Workbook.Worksheets[0];

            // Set and create Course
            string fileName = grades.Name;
            Course = CourseReader(rosterSheet, gradesSheet, fileName, rosterType);

            // Add students to course
            StudentReader(rosterSheet, gradesSheet, rosterType, createCerts);

            // If the certs aren't created create them
            if (createCerts == true)
            {
                CertificateCreator creator = new CertificateCreator(gradesSheet, certPath, Course, certificateType);
                // Add the cert to each student
                int numStudents = Course.Students.Count;
                for (int i = 0; i < numStudents; i++)
                {
                    string firstName = Course.Students[i].FirstName;
                    string lastName = Course.Students[i].LastName;
                    string fullName = $"{firstName} {lastName}";
                    string cert = MatchCert(fullName);
                    Course.Students[i].Certification = cert;
                }
            }
        }

        public void StudentReader(ExcelWorksheet rosterSheet, ExcelWorksheet gradesSheet, int rosterType, bool createCerts)
        {
            // Compare the data in each row
            int skip = 0; // If we need to skip when using the VA roster
            var rowCount = Math.Max(rosterSheet.Dimension.End.Row, gradesSheet.Dimension.End.Row);
            for (var row = 1; row <= rowCount; row++)
            {
                // Declare roster varibles
                var rosterFirstName = "";
                var rosterLastName = "";
                var rosterEmail = "";

                // BMRA roster
                if (rosterType == 1)
                {
                    int rosterSpacing = row + BMRAROSTERSPACE;
                    rosterFirstName = rosterSheet.Cells[rosterSpacing, 2].Value?.ToString();
                    rosterLastName = rosterSheet.Cells[rosterSpacing, 3].Value?.ToString();
                    rosterEmail = rosterSheet.Cells[rosterSpacing, 4].Value?.ToString();
                }
                // VA roster
                if (rosterType == 2)
                {
                    int rosterSpacing = row + GRADESPACE; // Use GRADESPACE since it's only skipping the header
                    var rosterPass = rosterSheet.Cells[rosterSpacing + skip, 7].Value?.ToString();
                    while (rosterPass == "N")
                    {
                        skip++;
                        rosterPass = rosterSheet.Cells[rosterSpacing + skip, 7].Value?.ToString();
                    }
                    rosterFirstName = rosterSheet.Cells[rosterSpacing + skip, 2].Value?.ToString();
                    rosterLastName = rosterSheet.Cells[rosterSpacing + skip, 1].Value?.ToString();
                    rosterEmail = rosterSheet.Cells[rosterSpacing + skip, 3].Value?.ToString();

                }

                int gradeSpacing = row + GRADESPACE;
                var gradesFirstName = gradesSheet.Cells[gradeSpacing, 3].Value?.ToString();
                var gradesLastName = gradesSheet.Cells[gradeSpacing, 5].Value?.ToString();
                var gradesGrade = gradesSheet.Cells[gradeSpacing, 6].Value?.ToString();

                string firstName = rosterFirstName ?? "";
                string lastName = rosterLastName ?? "";
                string email = rosterEmail ?? "";
                string grade = gradesGrade ?? "";

                // Case where certs need to be created
                string cert = "No Certificate Found";
                // If Certs are already created
                if (createCerts == false)
                {
                    string fullName = $"{firstName} {lastName}";
                    cert = MatchCert(fullName);
                }

                if (rosterFirstName == gradesFirstName && rosterLastName == gradesLastName && rosterFirstName != null && GradeCheck(grade))
                {
                    Student student = new Student(firstName, lastName, email, cert, grade);
                    Course.AddStudent(student);
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
            if (isDouble && result >= (double)minimumPassingGrade/100)
            {
                return true;
            }
            return false;
        }

        public Course CourseReader(ExcelWorksheet rosterSheet, ExcelWorksheet gradesSheet, string fileName, int rosterType)
        {
            // Varibles for data gathered from the course sheet
            var instructor = "";
/*            if (rosterType == 1)
            {
                *//*instructor = rosterSheet.Cells[7, 10].Value?.ToString();
                instructor = ColonSplit(instructor ?? "");*//*
                instructor = "";
            }
            if (rosterType == 2)
            {
                // MissingInfoForm missingInfoForm = new MissingInfoForm();
                // missingInfoForm.ShowDialog();
                //instructor = missingInfoForm.MissingData;
                instructor = ""; // Instructor doesn't really play a role rn so we can just skip this for simplicity
            }*/

            var courseName = gradesSheet.Cells[2, 7].Value?.ToString();
            var courseId = gradesSheet.Cells[2, 1].Value?.ToString();
            var agency = gradesSheet.Cells[2, 10].Value?.ToString();
            var location = gradesSheet.Cells[2, 11].Value?.ToString();
            var startDateString = gradesSheet.Cells[2, 8].Value?.ToString();
            var endDateString = gradesSheet.Cells[2, 9].Value?.ToString();
            DateTime startDate = StringDateTime(startDateString ?? "");
            DateTime endDate = StringDateTime(endDateString ?? "");

            // Get a possible courseId override
            string courseNamingScheme = fileName;
            courseNamingScheme = courseNamingScheme.Replace("BMRA Roster and Grades - ", "");
            courseNamingScheme = courseNamingScheme.Replace(".xlsx", "");

            if (courseId != courseNamingScheme)
            {
                courseId = courseNamingScheme;
            }

            return new Course(courseName ?? "", courseNamingScheme ?? "", courseId ?? "", agency ?? "", instructor ?? "", location ?? "", startDate, endDate);
        }


        private static DateTime StringDateTime(string date)
        {
            return DateTime.Parse(date);
        }

        private static string ColonSplit(string input)
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
                return "Could not find certificate for" + searchString;
            }
        }

        private static bool IsBMRARoster(ExcelPackage rosterExcel)
        {
            var worksheet = rosterExcel.Workbook.Worksheets[0]; // get the first worksheet
            var cell = worksheet.Cells["A1"]; // get the cell A1
            var value = cell.Value; // get the value of the cell A1
            string? a1Data = value.ToString();

            // Hardcoded the default Roster
            if (value != null && rosterExcel != null && a1Data != null)
            {
                if (a1Data.Contains("BUSINESS MANAGEMENT RESEARCH ASSOCIATES, INC"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

    }
}