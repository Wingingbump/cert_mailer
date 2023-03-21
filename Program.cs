using SendDraftEmail;

namespace cert_mailer
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
/*            string rosterPath = @"F:\work_project\TestEnv\11023.0001 SBA ICG-EL Virtual - NN\Attendance Roster - 11023.0001.xlsx";
            FileInfo rosterInfo = new FileInfo(rosterPath);
            string gradesPath = @"F:\work_project\TestEnv\11023.0001 SBA ICG-EL Virtual - NN\BMRA Roster and Grades - 11023.0001.xlsx";
            FileInfo gradesInfo = new FileInfo(gradesPath);
            string certPath = @"F:\work_project\TestEnv\11023.0001 SBA ICG-EL Virtual - NN\EOC\Certs";

            DataRead test = new DataRead(rosterInfo, gradesInfo, certPath);
            string courseName = test.course.CourseName;
            string courseId = test.course.CourseId;
            foreach (Student student in test.course.Students)
            {
                EmailBuilder message = new EmailBuilder(student.Email, courseName, courseId, student.Certification);
                message.CreateDraft();
            }
*/
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}
