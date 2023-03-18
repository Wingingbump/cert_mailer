using OfficeOpenXml;

namespace cert_mailer
{
    public class DataRead
    {
        const int GRADESPACE = 1;
        const int ROSTERSPACE = 11;
        private ExcelWorksheet? rosterSheet;
        private ExcelWorksheet? gradesSheet;
        private String certPath;
        public Course course { get; }

        public DataRead(FileInfo roster, FileInfo grades, String certPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var rosterExcel = new ExcelPackage(roster);
            // Get the Worksheet for the Roster
            foreach (ExcelWorksheet sheet in rosterExcel.Workbook.Worksheets)
            {
                if ((sheet.Name).Contains("EOC"))
                {
                    rosterSheet = sheet;
                    break;
                }
            }
            if (rosterSheet == null)
            {
                throw new Exception("The worksheet with name containing 'EOC' was not found.");
            }
            // Get the Worksheet for the Grades
            using var GradesExcel = new ExcelPackage(grades);
            gradesSheet = GradesExcel.Workbook.Worksheets[0];
            // Set CertPath
            this.certPath = certPath;
            // Set Course
            this.course = courseReader(rosterSheet, gradesSheet);
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
                string cert = matchCert(fullName);

                if (rosterFirstName == gradesFirstName && rosterLastName == gradesLastName && rosterFirstName != null)
                {
                    Student student = new Student(firstName, lastName, email, cert, grade);
                    //Console.WriteLine(student.ToString());
                    course.AddStudent(student);
                }
            }
        }

        public Course courseReader(ExcelWorksheet rosterSheet, ExcelWorksheet gradesSheet)
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

        private List<string> GetFileNames(string directoryPath)
        {
            List<string> fileNames = new List<string>();

            foreach (string filePath in Directory.GetFiles(directoryPath))
            {
                fileNames.Add(Path.GetFileName(filePath));
            }

            return fileNames;
        }

        private string matchCert(string searchString)
        {
            string[] files = Directory.GetFiles(certPath, "*.pdf");

            foreach (string file in files)
            {
                if (Path.GetFileNameWithoutExtension(file).Contains(searchString))
                {
                    return file;
                }
            }

            return "ERROR NOT FOUND";
        }

    }
}
