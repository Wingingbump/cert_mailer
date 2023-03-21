namespace cert_mailer
{
    public partial class Form1 : Form
    {

        private string? rosterPath = null;
        private string? gradesPath = null;
        private string? certPath = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(380, 300);
            this.MaximumSize = new Size(800, 800);
        }

        private void Browse1_Click(object sender, EventArgs e)
        {
            // Show the file explorer
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.Title = "Select a BMRA Roster Excel";

                // Update the "Roster" text box with the selected excel path
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    rosterPath = dialog.FileName;
                    Roster.Text = Path.GetFileName(dialog.FileName);
                }
            }
        }

        private void Browse2_Click(object sender, EventArgs e)
        {
            // Show the file explorer
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.Title = "Select a BMRA Roster and Grades Excel";

                // Update the "Grades" text box with the selected excel path
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    gradesPath = dialog.FileName;
                    Grades.Text = Path.GetFileName(dialog.FileName);
                }
            }
        }


        private void Browse3_Click(object sender, EventArgs e)
        {
            // Show the file explorer and get the selected directory
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a Certificate Directory";
                dialog.ShowNewFolderButton = false;
                DialogResult result = dialog.ShowDialog();

                // Update the "Cert" text box with the selected directory path
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    certPath = dialog.SelectedPath;
                    Certs.Text = Path.GetFileName(dialog.SelectedPath);
                    Certs.Text = certPath;
                }
            }
        }

        private void Submit_Click(object sender, EventArgs e)
        {
            if (rosterPath == null || certPath == null || gradesPath == null)
            {
                MessageBox.Show("Error: Not all fields filled", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    FileInfo rosterInfo = new FileInfo(rosterPath);
                    FileInfo gradesInfo = new FileInfo(gradesPath);

                    DataRead test = new DataRead(rosterInfo, gradesInfo, certPath);

                    string courseName = test.course.CourseName;
                    string courseId = test.course.CourseId;

                    foreach (Student student in test.course.Students)
                    {
                        EmailBuilder message = new EmailBuilder(student.Email, courseName, courseId, student.Certification);
                        message.CreateDraft();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred: " + ex.Message, "Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

    }
}