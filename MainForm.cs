using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ORSResearchTool
{
    public class MainForm : Form
    {
        // ── Controls ──────────────────────────────────────────────────────────────────
        private Label       lblInputFile;
        private TextBox     txtInputFile;
        private Button      btnBrowseInput;

        private Label       lblOutputFile;
        private TextBox     txtOutputFile;
        private Button      btnBrowseOutput;

        private GroupBox    grpConnection;
        private Label       lblHost;
        private TextBox     txtHost;
        private Label       lblLogin;
        private TextBox     txtLogin;
        private Label       lblPassword;
        private TextBox     txtPassword;

        private Label       lblDelayAction;
        private NumericUpDown nudActionDelay;
        private Label       lblDelayScroll;
        private NumericUpDown nudScrollDelay;

        private Button      btnStart;
        private Button      btnCancel;

        private ProgressBar progressBar;
        private Label       lblProgress;
        private RichTextBox rtbLog;

        private bool        _running = false;
        private System.Threading.CancellationTokenSource _cts;

        public MainForm()
        {
            InitializeComponent();
        }

        // ── UI Construction ───────────────────────────────────────────────────────────
        private void InitializeComponent()
        {
            this.Text            = "ORS Research Tool";
            this.Size            = new Size(700, 620);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox     = false;
            this.StartPosition   = FormStartPosition.CenterScreen;
            this.Font            = new Font("Arial", 9f);

            int y = 12;

            // ── Input file ────────────────────────────────────────────────────────────
            lblInputFile = new Label  { Text = "Input Excel File:", Left = 12, Top = y, Width = 120, AutoSize = false };
            txtInputFile = new TextBox { Left = 135, Top = y - 2, Width = 440, ReadOnly = true };
            btnBrowseInput = new Button { Text = "Browse…", Left = 582, Top = y - 2, Width = 80 };
            btnBrowseInput.Click += BtnBrowseInput_Click;
            y += 30;

            // ── Output file ───────────────────────────────────────────────────────────
            lblOutputFile = new Label  { Text = "Output Excel File:", Left = 12, Top = y, Width = 120, AutoSize = false };
            txtOutputFile = new TextBox { Left = 135, Top = y - 2, Width = 440 };
            btnBrowseOutput = new Button { Text = "Browse…", Left = 582, Top = y - 2, Width = 80 };
            btnBrowseOutput.Click += BtnBrowseOutput_Click;
            y += 36;

            // ── Connection group ──────────────────────────────────────────────────────
            grpConnection = new GroupBox { Text = "AS400 Connection", Left = 10, Top = y, Width = 660, Height = 110 };

            lblHost     = new Label  { Text = "Host:",     Left = 10, Top = 20, Width = 100, AutoSize = false, Parent = grpConnection };
            txtHost     = new TextBox { Left = 115, Top = 18, Width = 520, Parent = grpConnection };
            txtHost.Text = "YOUR_AS400_HOST"; // default – user must update

            lblLogin    = new Label  { Text = "Payloc Login:",  Left = 10, Top = 50, Width = 100, AutoSize = false, Parent = grpConnection };
            txtLogin    = new TextBox { Left = 115, Top = 48, Width = 200, Parent = grpConnection };

            lblPassword = new Label  { Text = "Password:",   Left = 330, Top = 50, Width = 80, AutoSize = false, Parent = grpConnection };
            txtPassword = new TextBox { Left = 415, Top = 48, Width = 220, UseSystemPasswordChar = true, Parent = grpConnection };

            lblDelayAction = new Label { Text = "Action delay (ms):", Left = 10, Top = 80, Width = 130, AutoSize = false, Parent = grpConnection };
            nudActionDelay = new NumericUpDown { Left = 145, Top = 78, Width = 70, Minimum = 100, Maximum = 5000, Value = 500, Parent = grpConnection };

            lblDelayScroll = new Label { Text = "Scroll delay (ms):", Left = 240, Top = 80, Width = 130, AutoSize = false, Parent = grpConnection };
            nudScrollDelay = new NumericUpDown { Left = 375, Top = 78, Width = 70, Minimum = 100, Maximum = 3000, Value = 300, Parent = grpConnection };

            y += grpConnection.Height + 10;

            // ── Buttons ───────────────────────────────────────────────────────────────
            btnStart  = new Button { Text = "▶  Start Research", Left = 12,  Top = y, Width = 160, Height = 30 };
            btnCancel = new Button { Text = "■  Cancel",          Left = 180, Top = y, Width = 100, Height = 30, Enabled = false };
            btnStart.Click  += BtnStart_Click;
            btnCancel.Click += BtnCancel_Click;
            y += 40;

            // ── Progress ──────────────────────────────────────────────────────────────
            progressBar = new ProgressBar { Left = 12, Top = y, Width = 660, Height = 18 };
            y += 24;
            lblProgress = new Label { Left = 12, Top = y, Width = 660, AutoSize = false, Text = "Ready." };
            y += 22;

            // ── Log ───────────────────────────────────────────────────────────────────
            rtbLog = new RichTextBox
            {
                Left       = 12,
                Top        = y,
                Width      = 660,
                Height     = 120,
                ReadOnly   = true,
                BackColor  = Color.Black,
                ForeColor  = Color.LimeGreen,
                Font       = new Font("Consolas", 8.5f),
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            this.Controls.AddRange(new Control[]
            {
                lblInputFile, txtInputFile, btnBrowseInput,
                lblOutputFile, txtOutputFile, btnBrowseOutput,
                grpConnection,
                btnStart, btnCancel,
                progressBar, lblProgress,
                rtbLog
            });
        }

        // ── Browse handlers ───────────────────────────────────────────────────────────

        private void BtnBrowseInput_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Title  = "Select Input Excel File",
                Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files|*.*"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtInputFile.Text = dlg.FileName;
                // Pre-fill output path
                if (string.IsNullOrWhiteSpace(txtOutputFile.Text))
                {
                    string dir  = Path.GetDirectoryName(dlg.FileName);
                    string name = Path.GetFileNameWithoutExtension(dlg.FileName);
                    txtOutputFile.Text = Path.Combine(dir, $"{name}_Results.xlsx");
                }
            }
        }

        private void BtnBrowseOutput_Click(object sender, EventArgs e)
        {
            using var dlg = new SaveFileDialog
            {
                Title  = "Save Output Excel File",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                txtOutputFile.Text = dlg.FileName;
        }

        // ── Start / Cancel ────────────────────────────────────────────────────────────

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            SetRunning(true);
            AppendLog("=== ORS Research Started ===");

            _cts = new System.Threading.CancellationTokenSource();

            string inputFile   = txtInputFile.Text;
            string outputFile  = txtOutputFile.Text;
            string host        = txtHost.Text.Trim();
            string login       = txtLogin.Text.Trim();
            string password    = txtPassword.Text;
            int    actionDelay = (int)nudActionDelay.Value;
            int    scrollDelay = (int)nudScrollDelay.Value;

            List<OrsRecord> results = null;

            try
            {
                // Read IDs on UI thread (fast)
                List<string> orsIds = ExcelHandler.ReadOrsIds(inputFile);
                AppendLog($"Loaded {orsIds.Count} ORS IDs from input file.");
                progressBar.Maximum = orsIds.Count;
                progressBar.Value   = 0;

                // Run session work on background thread
                results = await Task.Run(() =>
                {
                    using var session = new AS400Session(host, login, password)
                    {
                        ActionDelayMs    = actionDelay,
                        PageScrollDelayMs = scrollDelay
                    };

                    session.Connect();

                    var researcher = new OrsResearcher(session);

                    researcher.OnStatusUpdate += msg =>
                        this.BeginInvoke(new Action(() => AppendLog(msg)));

                    researcher.OnProgress += (cur, total) =>
                        this.BeginInvoke(new Action(() =>
                        {
                            progressBar.Value  = cur;
                            lblProgress.Text   = $"Processing {cur} of {total}…";
                        }));

                    return researcher.Research(orsIds);
                }, _cts.Token);

                // Write output
                ExcelHandler.WriteResults(outputFile, results);
                AppendLog($"=== Done! Output saved to: {outputFile} ===");
                lblProgress.Text = $"Completed. {results.Count} records written.";

                MessageBox.Show(
                    $"Research complete!\n\nRecords processed : {results.Count}\nOutput file       : {outputFile}",
                    "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                AppendLog("=== Research cancelled by user. ===");
                lblProgress.Text = "Cancelled.";
            }
            catch (Exception ex)
            {
                AppendLog($"FATAL ERROR: {ex.Message}");
                MessageBox.Show($"An error occurred:\n\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetRunning(false);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_running)
            {
                _cts?.Cancel();
                AppendLog("Cancellation requested – finishing current ORS…");
                btnCancel.Enabled = false;
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────────────

        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(txtInputFile.Text) || !File.Exists(txtInputFile.Text))
            {
                MessageBox.Show("Please select a valid input Excel file.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtOutputFile.Text))
            {
                MessageBox.Show("Please specify an output file path.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtHost.Text))
            {
                MessageBox.Show("Please enter the AS400 host address.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtLogin.Text) || string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Please enter Payloc login credentials.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private void SetRunning(bool running)
        {
            _running              = running;
            btnStart.Enabled      = !running;
            btnCancel.Enabled     = running;
            btnBrowseInput.Enabled  = !running;
            btnBrowseOutput.Enabled = !running;
            grpConnection.Enabled   = !running;
        }

        private void AppendLog(string message)
        {
            if (rtbLog.InvokeRequired)
            {
                rtbLog.BeginInvoke(new Action<string>(AppendLog), message);
                return;
            }
            rtbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            rtbLog.ScrollToCaret();
        }
    }
}
