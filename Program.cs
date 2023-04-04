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
            /*            string rosterPath = @"F:\tests\96523.0024 VAAA PM FPM 334-VA Virtual - JB\96523.0024 VAAA PM FPM 334-VA Virtual - JB\Attendance Roster - FPM 334 003, Mar. 23 - 24, 2023.xlsx";
                        //string rosterPath = @"F:\work_project\TestEnv\11023.0001 SBA ICG-EL Virtual - NN\Attendance Roster - 11023.0001.xlsx";
                        FileInfo rosterInfo = new FileInfo(rosterPath);
                        string gradesPath = @"F:\tests\96523.0024 VAAA PM FPM 334-VA Virtual - JB\96523.0024 VAAA PM FPM 334-VA Virtual - JB\BMRA Roster and Grades - FPM 334 003, Mar. 23 - 24, 2023.xlsx";
                        //string gradesPath = @"F:\work_project\TestEnv\11023.0001 SBA ICG-EL Virtual - NN\BMRA Roster and Grades - 11023.0001.xlsx";
                        FileInfo gradesInfo = new FileInfo(gradesPath);
                        string certPath = @"F:\tests\96523.0024 VAAA PM FPM 334-VA Virtual - JB\96523.0024 VAAA PM FPM 334-VA Virtual - JB\EOC\Certs";
                        //string certPath = @"F:\work_project\TestEnv\11023.0001 SBA ICG-EL Virtual - NN\EOC\Certs";

                        DataRead test = new DataRead(rosterInfo, gradesInfo, certPath);
                        string courseName = test.course.CourseName;
                        string courseId = test.course.CourseId;

                        foreach (Student student in test.course.Students)
                        {
                            EmailBuilder message = new EmailBuilder(student.Email, courseName, courseId, student.Certification);
                            message.CreateDraft();
                        }*/

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CertMailerForm());
        }
    }
}
