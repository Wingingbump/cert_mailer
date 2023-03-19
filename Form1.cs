namespace cert_mailer
{
    public partial class Form1 : Form
    {

        private string rosterPath = "";
        private string gradesPath = "";
        private string certsPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // What is this supposed to do?
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
                    certsPath = dialog.SelectedPath;
                    Certs.Text = Path.GetFileName(dialog.SelectedPath);
                    Certs.Text = certsPath;
                }
            }
        }

        private void Submit_Click(object sender, EventArgs e)
        {
            
        }
    }
}