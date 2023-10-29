using System;
using System.Windows.Forms;

namespace SejinTraceability
{
    public partial class ProjectSelectionForm : Form
    {
        private ComboBox ComboBoxProjects;
        private Button OkButton;

        public event EventHandler<string> ProjectSelected;

        public ProjectSelectionForm()
        {
            InitializeComponent();
            ComboBoxProjects.Items.AddRange(new object[] { "PM1", "PM2", "PM3" });
            ComboBoxProjects.SelectedIndex = 0; // Ustaw domyślnie wybrany element na pierwszy ("PM1")
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            // Tutaj możesz dodać logikę, która pobierze wybrany projekt z formularza.
            string selectedProject = ComboBoxProjects.SelectedItem?.ToString();

            // Upewnij się, że coś zostało wybrane
            if (!string.IsNullOrEmpty(selectedProject))
            {
                OnProjectSelected(selectedProject);
            }
            else
            {
                MessageBox.Show("Proszę wybrać projekt.");
            }
        }

        protected virtual void OnProjectSelected(string selectedProject)
        {
            ProjectSelected?.Invoke(this, selectedProject);
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close(); // Zamknij formularz po anulowaniu
        }

        private void InitializeComponent()
        {
            this.ComboBoxProjects = new System.Windows.Forms.ComboBox();
            this.OkButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ComboBoxProjects
            // 
            this.ComboBoxProjects.FormattingEnabled = true;
            this.ComboBoxProjects.Location = new System.Drawing.Point(33, 109);
            this.ComboBoxProjects.Name = "ComboBoxProjects";
            this.ComboBoxProjects.Size = new System.Drawing.Size(121, 23);
            this.ComboBoxProjects.TabIndex = 0;
            // 
            // OkButton
            // 
            this.OkButton.Location = new System.Drawing.Point(197, 109);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "button1";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // ProjectSelectionForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.ComboBoxProjects);
            this.Name = "ProjectSelectionForm";
            this.ResumeLayout(false);

        }
    }
}
