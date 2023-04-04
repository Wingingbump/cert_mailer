using OfficeOpenXml;

namespace cert_mailer
{
    public class DataRead
    {
        Dictionary<int, string> rosterType = new Dictionary<int, string>()
        {
            {1, "BMRA"},
            {2, "VA"},
            {3, "OTHER"}
        };
        const int GRADESPACE = 1;
        const int BMRAROSTERSPACE = 11;
        private Dictionary<string, string> certMap = new Dictionary<string, string>();
        private string certPath;
        public Course course { get; }

        public DataRead(FileInfo roster, FileInfo grades, String certPath)
        {
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
                rosterSheet = rosterExcel.Workbook.Worksheets.LastOrDefault();
                rosterType = 2;
            }

            if (rosterSheet == null)
            {
                throw new Exception("The worksheet was not found.");
            }


            // Get the Worksheet for the Grades
            using var gradesExcel = new ExcelPackage(grades);
            using var gradesSheet = gradesExcel.Workbook.Worksheets[0];
            // Set CertPath
            this.certPath = certPath;
            // Set Course
            string fileName = grades.Name;
            course = courseReader(rosterSheet, gradesSheet, fileName, rosterType);
            studentReader(rosterSheet, gradesSheet, rosterType);
        }

        public void studentReader(ExcelWorksheet RosterSheet, ExcelWorksheet GradesSheet, int rosterType)
        {
            // Compare the data in each row
            int skip = 0; // If we need to skip when using the VA roster
            var rowCount = Math.Max(RosterSheet.Dimension.End.Row, GradesSheet.Dimension.End.Row);
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
                    rosterFirstName = RosterSheet.Cells[rosterSpacing, 2].Value?.ToString();
                    rosterLastName = RosterSheet.Cells[rosterSpacing, 3].Value?.ToString();
                    rosterEmail = RosterSheet.Cells[rosterSpacing, 4].Value?.ToString();
                }
                // VA roster
                if (rosterType == 2) 
                {
                    int rosterSpacing = row + GRADESPACE; // Use GRADESPACE since it's only skipping the header
                    var rosterPass = RosterSheet.Cells[rosterSpacing, 7].Value?.ToString();
                    if (rosterPass == "N")
                    {
                        skip++;
                    }
                    rosterFirstName = RosterSheet.Cells[rosterSpacing + skip, 2].Value?.ToString();
                    rosterLastName = RosterSheet.Cells[rosterSpacing + skip, 1].Value?.ToString();
                    rosterEmail = RosterSheet.Cells[rosterSpacing + skip, 3].Value?.ToString();

                }



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

        public Course courseReader(ExcelWorksheet rosterSheet, ExcelWorksheet gradesSheet, string fileName, int rosterType)
        {
            // Varibles for data gathered from the course sheet
            var instructor = "";
            if (rosterType == 1)
            {
                instructor = rosterSheet.Cells[7, 10].Value?.ToString();
                instructor = colonSplit(instructor ?? "");
            }
            if (rosterType == 2)
            {
                MissingInfoForm missingInfoForm = new MissingInfoForm();
                missingInfoForm.ShowDialog();
                instructor = missingInfoForm.MissingData;
            }

            var courseName = gradesSheet.Cells[2, 7].Value?.ToString();
            var courseId = gradesSheet.Cells[2, 1].Value?.ToString();
            var agency = gradesSheet.Cells[2, 10].Value?.ToString();
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

        private bool IsBMRARoster(ExcelPackage rosterExcel)
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
                    Console.WriteLine("TRUE");
                    Console.WriteLine(a1Data);
                    return true;
                }
                else
                {
                    Console.WriteLine("FALSE");
                    Console.WriteLine(a1Data);
                    return false;
                }
            }
            return false;
        }

    }
}