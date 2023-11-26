using System;
using System.Windows.Forms;

namespace SejinTraceability
{
    public partial class ProjectSelectionForm : Form
    {
        private ComboBox ComboBoxProjects;
        private Button OkButton;
        private new Button CancelButton;

        public event EventHandler<string> ProjectSelected;
        private bool projectSelected = false;
        public event EventHandler<string> ProjectSelectedOnce;
        public ProjectSelectionForm()
        {
            InitializeComponent();
            ComboBoxProjects.Items.Add("Projekt 1");
            ComboBoxProjects.Items.Add("Projekt 2");
            ComboBoxProjects.Items.Add("Projekt 3");
        }
        private void OnProjectSelected(string selectedProject)
        {
            projectSelected = true;
            ProjectSelected?.Invoke(this, selectedProject);
            Close();
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

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close(); // Zamknij formularz po anulowaniu
        }

   
        private void InitializeComponent()
        {
            this.ComboBoxProjects = new System.Windows.Forms.ComboBox();
            this.OkButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ComboBoxProjects
            // 
            this.ComboBoxProjects.FormattingEnabled = true;
            this.ComboBoxProjects.Location = new System.Drawing.Point(12, 12);
            this.ComboBoxProjects.Name = "ComboBoxProjects";
            this.ComboBoxProjects.Size = new System.Drawing.Size(121, 23);
            this.ComboBoxProjects.TabIndex = 0;
            // 
            // OkButton
            // 
            this.OkButton.Location = new System.Drawing.Point(12, 41);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(47, 23);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "OK";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(65, 41);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(68, 23);
            this.CancelButton.TabIndex = 2;
            this.CancelButton.Text = "Anuluj";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ProjectSelectionForm
            // 
            this.ClientSize = new System.Drawing.Size(147, 79);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.ComboBoxProjects);
            this.Name = "ProjectSelectionForm";
            this.ResumeLayout(false);

        }
    }
}
