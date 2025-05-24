using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace ContactManager
{
    public partial class Form1 : Form
    {
        // Modern color palette
        private Color primaryColor = Color.FromArgb(68, 114, 196);  // Soft blue
        private Color secondaryColor = Color.FromArgb(240, 240, 240); // Light gray
        private Color accentColor = Color.FromArgb(0, 150, 136);    // Teal
        private Color textColor = Color.FromArgb(51, 51, 51);       // Dark gray
        private Color cardColor = Color.White;

        // Database connection
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ContactManager.accdb;";

        // Form controls
        private TextBox txtFirstName, txtLastName, txtEmail, txtPhone, txtSearch;
        private DataGridView dgvContacts;
        private Button btnAdd, btnUpdate, btnDelete, btnClear, btnExport, btnRefresh, btnSearch;

        // For rounded corners
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            ApplyModernTheme();
            LoadContacts();
        }

        private void InitializeCustomComponents()
        {
            // Form settings
            this.Text = "Modern Contact Manager";
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 10);
            this.BackColor = secondaryColor;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // Header Panel
            Panel headerPanel = new Panel
            {
                BackColor = primaryColor,
                Dock = DockStyle.Top,
                Height = 80
            };

            // Title Label
            Label lblTitle = new Label
            {
                Text = "Contact Manager",
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(30, 20)
            };

            // Main Content Panel
            Panel contentPanel = new Panel
            {
                BackColor = secondaryColor,
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };

            // Card Panel for form controls
            Panel formCard = new Panel
            {
                BackColor = cardColor,
                Size = new Size(840, 200),
                Location = new Point(20, 20 + headerPanel.Height) // <-- Move below header
            };

            // Search Panel
            Panel searchPanel = new Panel
            {
                BackColor = cardColor,
                Size = new Size(840, 60),
                Location = new Point(20, 240 + headerPanel.Height) // <-- Move below header
            };

            // Grid Panel
            Panel gridPanel = new Panel
            {
                BackColor = cardColor,
                Size = new Size(840, 300),
                Location = new Point(20, 320 + headerPanel.Height) // <-- Move below header
            };

            // Add controls to panels
            headerPanel.Controls.Add(lblTitle);
            
            // Form Card Controls
            AddFormControls(formCard);
            
            // Search Panel Controls
            AddSearchControls(searchPanel);
            
            // Grid Panel Controls
            AddGridControls(gridPanel);

            // Add panels to content panel
            contentPanel.Controls.Add(formCard);
            contentPanel.Controls.Add(searchPanel);
            contentPanel.Controls.Add(gridPanel);

            // Add panels to form
            this.Controls.Add(headerPanel);
            this.Controls.Add(contentPanel);
        }

        private void AddFormControls(Panel panel)
        {
            // Labels
            Label lblFirstName = new Label { Text = "First Name", Location = new Point(20, 20), Width = 100, ForeColor = textColor };
            Label lblLastName = new Label { Text = "Last Name", Location = new Point(20, 60), Width = 100, ForeColor = textColor };
            Label lblEmail = new Label { Text = "Email", Location = new Point(20, 100), Width = 100, ForeColor = textColor };
            Label lblPhone = new Label { Text = "Phone", Location = new Point(20, 140), Width = 100, ForeColor = textColor };

            // TextBoxes
            txtFirstName = CreateModernTextBox(new Point(140, 20), 250);
            txtLastName = CreateModernTextBox(new Point(140, 60), 250);
            txtEmail = CreateModernTextBox(new Point(140, 100), 250);
            txtPhone = CreateModernTextBox(new Point(140, 140), 250);

            // Buttons
            btnAdd = CreateModernButton("Add", new Point(420, 20), accentColor);
            btnUpdate = CreateModernButton("Update", new Point(420, 70), accentColor);
            btnDelete = CreateModernButton("Delete", new Point(420, 120), Color.FromArgb(229, 57, 53));
            btnClear = CreateModernButton("Clear", new Point(550, 20), secondaryColor);
            btnExport = CreateModernButton("Export", new Point(550, 70), primaryColor);
            btnRefresh = CreateModernButton("Refresh", new Point(550, 120), primaryColor);

            // Event handlers
            btnAdd.Click += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click += BtnClear_Click;
            btnExport.Click += BtnExport_Click;
            btnRefresh.Click += BtnRefresh_Click;

            // Add controls to panel
            panel.Controls.AddRange(new Control[] {
                lblFirstName, lblLastName, lblEmail, lblPhone,
                txtFirstName, txtLastName, txtEmail, txtPhone,
                btnAdd, btnUpdate, btnDelete, btnClear, btnExport, btnRefresh
            });
        }

        private void AddSearchControls(Panel panel)
        {
            txtSearch = CreateModernTextBox(new Point(20, 15), 600);
            btnSearch = CreateModernButton("Search", new Point(640, 15), primaryColor);

            txtSearch.PlaceholderText = "Search contacts...";
            txtSearch.TextChanged += TxtSearch_TextChanged;
            btnSearch.Click += BtnSearch_Click;

            panel.Controls.Add(txtSearch);
            panel.Controls.Add(btnSearch);
        }

        private void AddGridControls(Panel panel)
        {
            dgvContacts = new DataGridView
            {
                Location = new Point(20, 20),
                Size = new Size(800, 260),
                BackgroundColor = cardColor,
                BorderStyle = BorderStyle.None,
                EnableHeadersVisualStyles = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                ReadOnly = true
            };

            // Style the grid
            dgvContacts.DefaultCellStyle.Font = new Font("Segoe UI", 9);
            dgvContacts.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            dgvContacts.ColumnHeadersDefaultCellStyle.BackColor = primaryColor;
            dgvContacts.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvContacts.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 248, 248);

            dgvContacts.SelectionChanged += DgvContacts_SelectionChanged;

            panel.Controls.Add(dgvContacts);
        }

        private TextBox CreateModernTextBox(Point location, int width)
        {
            return new TextBox
            {
                Location = location,
                Width = width,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                ForeColor = textColor,
                Font = new Font("Segoe UI", 10)
            };
        }

        private Button CreateModernButton(string text, Point location, Color backColor)
        {
            Button button = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(120, 40),
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };

            button.FlatAppearance.BorderSize = 0;
            button.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, button.Width, button.Height, 15, 15));

            // Hover effects
            button.MouseEnter += (sender, e) => 
            {
                button.BackColor = ControlPaint.Light(backColor, 0.1f);
            };
            button.MouseLeave += (sender, e) => 
            {
                button.BackColor = backColor;
            };

            return button;
        }

        private void ApplyModernTheme()
        {
            // Apply rounded corners to all panels
            foreach (Control control in this.Controls)
            {
                if (control is Panel panel)
                {
                    foreach (Control panelControl in panel.Controls)
                    {
                        if (panelControl is Panel subPanel)
                        {
                            subPanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, subPanel.Width, subPanel.Height, 10, 10));
                            subPanel.Padding = new Padding(10);
                        }
                    }
                }
            }
        }

        // Database Methods
        private void LoadContacts(string searchTerm = "")
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string query = "SELECT ID, FirstName, LastName, Email, Phone FROM Contacts";
                    
                    if (!string.IsNullOrEmpty(searchTerm))
                    {
                        query += " WHERE FirstName LIKE ? OR LastName LIKE ? OR Email LIKE ? OR Phone LIKE ?";
                    }

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        if (!string.IsNullOrEmpty(searchTerm))
                        {
                            string likeTerm = $"%{searchTerm}%";
                            cmd.Parameters.AddWithValue("@p1", likeTerm);
                            cmd.Parameters.AddWithValue("@p2", likeTerm);
                            cmd.Parameters.AddWithValue("@p3", likeTerm);
                            cmd.Parameters.AddWithValue("@p4", likeTerm);
                        }

                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            dgvContacts.DataSource = dt;

                            // Configure columns
                            if (dgvContacts.Columns.Count > 0)
                            {
                                dgvContacts.Columns["ID"].Visible = false;
                                dgvContacts.Columns["FirstName"].HeaderText = "First Name";
                                dgvContacts.Columns["LastName"].HeaderText = "Last Name";
                                dgvContacts.Columns["Email"].HeaderText = "Email";
                                dgvContacts.Columns["Phone"].HeaderText = "Phone";
                                dgvContacts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("Error loading contacts: " + ex.Message);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFirstName.Text))
            {
                ShowError("First name is required");
                return;
            }

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string query = @"INSERT INTO Contacts (FirstName, LastName, Email, Phone) 
                                    VALUES (?, ?, ?, ?)";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@FirstName", txtFirstName.Text);
                        cmd.Parameters.AddWithValue("@LastName", txtLastName.Text);
                        cmd.Parameters.AddWithValue("@Email", txtEmail.Text);
                        cmd.Parameters.AddWithValue("@Phone", txtPhone.Text);

                        int rowsAffected = cmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            ShowSuccess("Contact added successfully!");
                            ClearFields();
                            LoadContacts();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("Error adding contact: " + ex.Message);
            }
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (dgvContacts.SelectedRows.Count == 0)
            {
                ShowError("Please select a contact to update");
                return;
            }

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string query = @"UPDATE Contacts SET 
                                    FirstName = ?, 
                                    LastName = ?, 
                                    Email = ?, 
                                    Phone = ? 
                                    WHERE ID = ?";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@FirstName", txtFirstName.Text);
                        cmd.Parameters.AddWithValue("@LastName", txtLastName.Text);
                        cmd.Parameters.AddWithValue("@Email", txtEmail.Text);
                        cmd.Parameters.AddWithValue("@Phone", txtPhone.Text);
                        cmd.Parameters.AddWithValue("@ID", dgvContacts.SelectedRows[0].Cells["ID"].Value);

                        int rowsAffected = cmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            ShowSuccess("Contact updated successfully!");
                            LoadContacts();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("Error updating contact: " + ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgvContacts.SelectedRows.Count == 0)
            {
                ShowError("Please select a contact to delete");
                return;
            }

            var confirm = MessageBox.Show("Are you sure you want to delete this contact?",
                                        "Confirm Delete",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string query = "DELETE FROM Contacts WHERE ID = ?";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", dgvContacts.SelectedRows[0].Cells["ID"].Value);

                        int rowsAffected = cmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            ShowSuccess("Contact deleted successfully!");
                            ClearFields();
                            LoadContacts();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("Error deleting contact: " + ex.Message);
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            ClearFields();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (dgvContacts.Rows.Count == 0)
            {
                ShowError("No contacts to export");
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "CSV file (*.csv)|*.csv", ValidateNames = true })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (StreamWriter sw = new StreamWriter(sfd.FileName))
                        {
                            // Write headers
                            for (int i = 0; i < dgvContacts.Columns.Count; i++)
                            {
                                if (dgvContacts.Columns[i].Visible)
                                {
                                    sw.Write(dgvContacts.Columns[i].HeaderText);
                                    if (i < dgvContacts.Columns.Count - 1) sw.Write(",");
                                }
                            }
                            sw.WriteLine();

                            // Write rows
                            foreach (DataGridViewRow row in dgvContacts.Rows)
                            {
                                for (int i = 0; i < dgvContacts.Columns.Count; i++)
                                {
                                    if (dgvContacts.Columns[i].Visible)
                                    {
                                        if (row.Cells[i].Value != null)
                                        {
                                            sw.Write("\"" + row.Cells[i].Value.ToString().Replace("\"", "\"\"") + "\"");
                                        }
                                        if (i < dgvContacts.Columns.Count - 1) sw.Write(",");
                                    }
                                }
                                sw.WriteLine();
                            }
                        }

                        ShowSuccess($"Contacts exported to {sfd.FileName}");
                    }
                    catch (Exception ex)
                    {
                        ShowError("Error exporting contacts: " + ex.Message);
                    }
                }
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadContacts();
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            LoadContacts(txtSearch.Text);
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            LoadContacts(txtSearch.Text);
        }

        private void DgvContacts_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvContacts.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dgvContacts.SelectedRows[0];
                txtFirstName.Text = row.Cells["FirstName"].Value?.ToString() ?? "";
                txtLastName.Text = row.Cells["LastName"].Value?.ToString() ?? "";
                txtEmail.Text = row.Cells["Email"].Value?.ToString() ?? "";
                txtPhone.Text = row.Cells["Phone"].Value?.ToString() ?? "";
            }
        }

        private void ClearFields()
        {
            txtFirstName.Text = "";
            txtLastName.Text = "";
            txtEmail.Text = "";
            txtPhone.Text = "";
            dgvContacts.ClearSelection();
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowSuccess(string message)
        {
            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}