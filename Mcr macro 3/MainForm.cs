using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ORSResearchTool.Database;
using ORSResearchTool.Models;

namespace ORSResearchTool
{
    public class MainForm : Form
    {
        // ── Controls ──────────────────────────────────────────────────────────
        private Label         lblInputFile;
        private TextBox       txtInputFile;
        private Button        btnBrowseInput;

        private Label         lblOutputFile;
        private TextBox       txtOutputFile;
        private Button        btnBrowseOutput;

        private GroupBox      grpConnection;
        private Label         lblSession;
        private TextBox       txtSession;
        private Label         lblSession2;
        private TextBox       txtSession2;
        private Label         lblLocality;
        private ComboBox      cmbLocality;          // ← Locality dropdown
        private Label         lblCredStatus;        // shows credential fetch result

        private Label         lblDelayAction;
        private NumericUpDown nudActionDelay;
        private Label         lblDelayScroll;
        private NumericUpDown nudScrollDelay;

        private Button        btnLoadCredentials;   // explicit credential prefetch
        private Button        btnStart;
        private Button        btnCancel;

        private ProgressBar   progressBar;
        private Label         lblProgress;
        private RichTextBox   rtbLog;

        // ── State ─────────────────────────────────────────────────────────────
        private bool                                _running;
        private CancellationTokenSource             _cts;
        private Dictionary<string, DroidCredential> _credCache;  // loaded from DB once

        public MainForm()
        {
            InitializeComponent();
        }

        // ── UI construction ───────────────────────────────────────────────────
        // NOTE: No inline  Parent = X  assignments anywhere — the WinForms
        // Designer parser cannot resolve them and throws NullReferenceException.
        // All parent/child relationships are established via Controls.Add() at
        // the bottom of this method.
        private void InitializeComponent()
        {
            // ── Instantiate every control first (no Parent= here) ─────────────
            lblInputFile       = new Label();
            txtInputFile       = new TextBox();
            btnBrowseInput     = new Button();

            lblOutputFile      = new Label();
            txtOutputFile      = new TextBox();
            btnBrowseOutput    = new Button();

            grpConnection      = new GroupBox();
            lblSession         = new Label();
            txtSession         = new TextBox();
            lblSession2        = new Label();
            txtSession2        = new TextBox();
            lblLocality        = new Label();
            cmbLocality        = new ComboBox();
            lblCredStatus      = new Label();
            lblDelayAction     = new Label();
            nudActionDelay     = new NumericUpDown();
            lblDelayScroll     = new Label();
            nudScrollDelay     = new NumericUpDown();
            btnLoadCredentials = new Button();

            btnStart           = new Button();
            btnCancel          = new Button();
            progressBar        = new ProgressBar();
            lblProgress        = new Label();
            rtbLog             = new RichTextBox();

            // Suspend layout while we configure everything
            SuspendLayout();
            grpConnection.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)nudActionDelay).BeginInit();
            ((System.ComponentModel.ISupportInitialize)nudScrollDelay).BeginInit();

            // ── Form ──────────────────────────────────────────────────────────
            Text            = "ORS Research Tool";
            ClientSize      = new Size(694, 620);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            StartPosition   = FormStartPosition.CenterScreen;
            Font            = new Font("Arial", 9f);

            // ── Input file row  (y=12) ────────────────────────────────────────
            lblInputFile.Text      = "Input Excel File:";
            lblInputFile.Left      = 12;
            lblInputFile.Top       = 14;
            lblInputFile.Width     = 120;
            lblInputFile.AutoSize  = false;

            txtInputFile.Left      = 135;
            txtInputFile.Top       = 12;
            txtInputFile.Width     = 440;
            txtInputFile.ReadOnly  = true;

            btnBrowseInput.Text    = "Browse…";
            btnBrowseInput.Left    = 582;
            btnBrowseInput.Top     = 11;
            btnBrowseInput.Width   = 80;
            btnBrowseInput.Click  += BtnBrowseInput_Click;

            // ── Output file row  (y=42) ───────────────────────────────────────
            lblOutputFile.Text     = "Output Excel File:";
            lblOutputFile.Left     = 12;
            lblOutputFile.Top      = 44;
            lblOutputFile.Width    = 120;
            lblOutputFile.AutoSize = false;

            txtOutputFile.Left     = 135;
            txtOutputFile.Top      = 42;
            txtOutputFile.Width    = 440;

            btnBrowseOutput.Text   = "Browse…";
            btnBrowseOutput.Left   = 582;
            btnBrowseOutput.Top    = 41;
            btnBrowseOutput.Width  = 80;
            btnBrowseOutput.Click += BtnBrowseOutput_Click;

            // ── GroupBox  (y=78) ──────────────────────────────────────────────
            grpConnection.Text   = "Session & Credentials (from DB)";
            grpConnection.Left   = 10;
            grpConnection.Top    = 78;
            grpConnection.Width  = 672;
            grpConnection.Height = 148;

            // Session name — inside GroupBox (coords relative to GroupBox)
            lblSession.Text      = "Session Name:";
            lblSession.Left      = 10;
            lblSession.Top       = 24;
            lblSession.Width     = 110;
            lblSession.AutoSize  = false;

            txtSession.Left      = 125;
            txtSession.Top       = 22;
            txtSession.Width     = 55;
            txtSession.Text      = "A";

            lblSession2.Text     = "Session B:";
            lblSession2.Left     = 190;
            lblSession2.Top      = 24;
            lblSession2.Width    = 75;
            lblSession2.AutoSize = false;

            txtSession2.Left = 268;
            txtSession2.Top  = 22;
            txtSession2.Width = 55;
            txtSession2.Text  = "B";

            // Locality label + combo — shifted right to accommodate Session B
            lblLocality.Text     = "Locality:";
            lblLocality.Left     = 338;
            lblLocality.Top      = 24;
            lblLocality.Width    = 60;
            lblLocality.AutoSize = false;

            cmbLocality.Left          = 400;
            cmbLocality.Top           = 21;
            cmbLocality.Width         = 160;
            cmbLocality.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLocality.Items.Add("Domestic");
            cmbLocality.Items.Add("Global");
            cmbLocality.SelectedIndex = 0;
            cmbLocality.SelectedIndexChanged += CmbLocality_Changed;

            // Credential status label
            lblCredStatus.Text      = "Credentials not loaded.";
            lblCredStatus.Left      = 10;
            lblCredStatus.Top       = 52;
            lblCredStatus.Width     = 645;
            lblCredStatus.Height    = 18;
            lblCredStatus.AutoSize  = false;
            lblCredStatus.ForeColor = Color.Gray;

            // Action delay
            lblDelayAction.Text     = "Action delay (ms):";
            lblDelayAction.Left     = 10;
            lblDelayAction.Top      = 80;
            lblDelayAction.Width    = 130;
            lblDelayAction.AutoSize = false;

            nudActionDelay.Left    = 145;
            nudActionDelay.Top     = 78;
            nudActionDelay.Width   = 70;
            nudActionDelay.Minimum = 100;
            nudActionDelay.Maximum = 5000;
            nudActionDelay.Value   = 500;

            // Scroll delay
            lblDelayScroll.Text     = "Scroll delay (ms):";
            lblDelayScroll.Left     = 245;
            lblDelayScroll.Top      = 80;
            lblDelayScroll.Width    = 130;
            lblDelayScroll.AutoSize = false;

            nudScrollDelay.Left    = 380;
            nudScrollDelay.Top     = 78;
            nudScrollDelay.Width   = 70;
            nudScrollDelay.Minimum = 100;
            nudScrollDelay.Maximum = 3000;
            nudScrollDelay.Value   = 300;

            // Load credentials button
            btnLoadCredentials.Text   = "Load Credentials from DB";
            btnLoadCredentials.Left   = 10;
            btnLoadCredentials.Top    = 110;
            btnLoadCredentials.Width  = 200;
            btnLoadCredentials.Height = 26;
            btnLoadCredentials.Click += BtnLoadCredentials_Click;

            // ── Action buttons  (y=236) ───────────────────────────────────────
            btnStart.Text    = "▶  Start Research";
            btnStart.Left    = 12;
            btnStart.Top     = 236;
            btnStart.Width   = 165;
            btnStart.Height  = 30;
            btnStart.Click  += BtnStart_Click;

            btnCancel.Text    = "■  Cancel";
            btnCancel.Left    = 185;
            btnCancel.Top     = 236;
            btnCancel.Width   = 100;
            btnCancel.Height  = 30;
            btnCancel.Enabled = false;
            btnCancel.Click  += BtnCancel_Click;

            // ── Progress  (y=276) ─────────────────────────────────────────────
            progressBar.Left   = 12;
            progressBar.Top    = 276;
            progressBar.Width  = 670;
            progressBar.Height = 18;

            lblProgress.Left     = 12;
            lblProgress.Top      = 298;
            lblProgress.Width    = 670;
            lblProgress.AutoSize = false;
            lblProgress.Text     = "Ready.";

            // ── Log  (y=322) ──────────────────────────────────────────────────
            rtbLog.Left       = 12;
            rtbLog.Top        = 322;
            rtbLog.Width      = 670;
            rtbLog.Height     = 268;
            rtbLog.ReadOnly   = true;
            rtbLog.BackColor  = Color.Black;
            rtbLog.ForeColor  = Color.LimeGreen;
            rtbLog.Font       = new Font("Consolas", 8.5f);
            rtbLog.ScrollBars = RichTextBoxScrollBars.Vertical;

            // ── Wire up parent/child via Controls.Add (Designer-safe) ─────────
            grpConnection.Controls.Add(lblSession);
            grpConnection.Controls.Add(txtSession);
            grpConnection.Controls.Add(lblSession2);
            grpConnection.Controls.Add(txtSession2);
            grpConnection.Controls.Add(lblLocality);
            grpConnection.Controls.Add(cmbLocality);
            grpConnection.Controls.Add(lblCredStatus);
            grpConnection.Controls.Add(lblDelayAction);
            grpConnection.Controls.Add(nudActionDelay);
            grpConnection.Controls.Add(lblDelayScroll);
            grpConnection.Controls.Add(nudScrollDelay);
            grpConnection.Controls.Add(btnLoadCredentials);

            Controls.Add(lblInputFile);
            Controls.Add(txtInputFile);
            Controls.Add(btnBrowseInput);
            Controls.Add(lblOutputFile);
            Controls.Add(txtOutputFile);
            Controls.Add(btnBrowseOutput);
            Controls.Add(grpConnection);
            Controls.Add(btnStart);
            Controls.Add(btnCancel);
            Controls.Add(progressBar);
            Controls.Add(lblProgress);
            Controls.Add(rtbLog);

            // Resume layout
            ((System.ComponentModel.ISupportInitialize)nudActionDelay).EndInit();
            ((System.ComponentModel.ISupportInitialize)nudScrollDelay).EndInit();
            grpConnection.ResumeLayout(false);
            ResumeLayout(false);
        }

        // ── Browse ────────────────────────────────────────────────────────────

        private void BtnBrowseInput_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog
            {
                Title  = "Select Input Excel File",
                Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files|*.*"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                txtInputFile.Text = dlg.FileName;
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
            using (var dlg = new SaveFileDialog
            {
                Title      = "Save Output Excel File",
                Filter     = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx"
            })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                    txtOutputFile.Text = dlg.FileName;
            }
        }

        // ── Locality dropdown changed ─────────────────────────────────────────

        private void CmbLocality_Changed(object sender, EventArgs e)
        {
            // If the cache is already loaded, validate the selected locality exists
            if (_credCache != null)
                UpdateCredentialStatusLabel();
        }

        // ── Load credentials ──────────────────────────────────────────────────

        private async void BtnLoadCredentials_Click(object sender, EventArgs e)
        {
            btnLoadCredentials.Enabled = false;
            lblCredStatus.ForeColor    = Color.DarkGoldenrod;
            lblCredStatus.Text         = "Connecting to SecondaryDB (Wipro_Corrected_Droid)…";

            try
            {
                _credCache = await Task.Run(() =>
                {
                    var db = new DatabaseHelper();
                    return db.FetchAllCredentials();
                });

                AppendLog($"Credentials loaded: {_credCache.Count} localities found.");
                UpdateCredentialStatusLabel();
            }
            catch (Exception ex)
            {
                lblCredStatus.ForeColor = Color.Red;
                lblCredStatus.Text      = "Failed to load credentials: " + ex.Message;
                AppendLog("ERROR loading credentials: " + ex.Message);
            }
            finally
            {
                btnLoadCredentials.Enabled = true;
            }
        }

        private void UpdateCredentialStatusLabel()
        {
            if (_credCache == null) return;

            string locality = cmbLocality.SelectedItem?.ToString() ?? string.Empty;

            if (_credCache.ContainsKey(locality))
            {
                var cred = _credCache[locality];
                lblCredStatus.ForeColor = Color.DarkGreen;
                lblCredStatus.Text      =
                    $"✔  Credential loaded for '{locality}'  —  " +
                    $"Payloc ID: {cred.ID}  |  {_credCache.Count} locality records in cache.";
            }
            else
            {
                lblCredStatus.ForeColor = Color.Red;
                lblCredStatus.Text      =
                    $"✘  No credential found for locality '{locality}' in DroidCred table.";
            }
        }

        // ── Start / Cancel ────────────────────────────────────────────────────

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            string locality    = cmbLocality.SelectedItem?.ToString() ?? string.Empty;
            var    cred        = _credCache[locality];
            string outputPath  = txtOutputFile.Text;
            string sessionName = txtSession.Text.Trim().ToUpper();
            int    actionDelay = (int)nudActionDelay.Value;
            int    scrollDelay = (int)nudScrollDelay.Value;

            SetRunning(true);
            AppendLog($"=== ORS Research Started — Locality: {locality} ===");
            AppendLog($"Payloc ID: {cred.ID}  Session: {sessionName}");

            _cts = new CancellationTokenSource();
            var ct = _cts.Token;

            List<OrsRecord> results = null;

            try
            {
                // ── Step 1: read ORS IDs (fast, UI thread) ───────────────────
                var orsIds = ExcelHandler.ReadOrsIds(txtInputFile.Text);
                AppendLog($"Loaded {orsIds.Count} ORS IDs from input file.");
                progressBar.Maximum = orsIds.Count;
                progressBar.Value   = 0;

                // ── Step 2: create the output file with header row NOW ────────
                // From this point, even if the process is cancelled or crashes,
                // the file exists and contains everything processed so far.
                ExcelHandler.InitialiseOutputFile(outputPath);
                AppendLog($"Output file created: {outputPath}");

                // ── Step 3: process on background thread ──────────────────────
                string sessionName2 = txtSession2.Text.Trim().ToUpper();

                results = await Task.Run(() =>
                {
                    // Session A — Payloc login, ORS navigation
                    var session = new AS400Session(sessionName)
                    {
                        ActionDelayMs     = actionDelay,
                        PageScrollDelayMs = scrollDelay
                    };
                    session.SetCredential(cred, locality);
                    session.Connect();

                    // Session B — Seamless login, MHI lookups
                    var mhi = new MhiSession(cred, sessionName2);
                    mhi.Connect();

                    var researcher = new OrsResearcher(session, mhi);

                    researcher.OnStatusUpdate += msg =>
                        BeginInvoke(new Action(() => AppendLog(msg)));

                    researcher.OnProgress += (cur, tot) =>
                        BeginInvoke(new Action(() =>
                        {
                            progressBar.Value = cur;
                            lblProgress.Text  = $"Processing {cur} of {tot}…";
                        }));

                    var res = researcher.Research(orsIds, outputPath, ct);
                    session.Dispose();
                    mhi.Dispose();
                    return res;

                }, ct);

                // ── Step 4: finalise — append summary block + auto-fit ────────
                ExcelHandler.FinaliseOutputFile(outputPath, results);

                bool cancelled = ct.IsCancellationRequested;
                string doneMsg = cancelled
                    ? $"Cancelled — {results.Count} records saved."
                    : $"Completed — {results.Count} records saved.";

                AppendLog($"=== {doneMsg} Output → {outputPath} ===");
                lblProgress.Text = doneMsg;

                MessageBox.Show(
                    $"{doneMsg}\n\nLocality    : {locality}\nOutput file : {outputPath}",
                    cancelled ? "Cancelled" : "Done",
                    MessageBoxButtons.OK,
                    cancelled ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                // FinaliseOutputFile even on hard cancel so the file is clean
                if (results != null)
                    TryFinalise(outputPath, results);

                AppendLog("=== Cancelled. Partial results saved to output file. ===");
                lblProgress.Text = "Cancelled — partial results saved.";

                MessageBox.Show(
                    $"Processing was cancelled.\nPartial results have been saved to:\n{outputPath}",
                    "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                // Even on a fatal crash, try to finalise whatever was checkpointed
                if (results != null)
                    TryFinalise(outputPath, results);

                AppendLog($"FATAL ERROR: {ex.Message}");

                if (ex.Message.Contains("Payloc login failed"))
                    MessageBox.Show("Payloc login failed.\nUNET system may not be available.",
                        "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show(
                        $"An error occurred:\n\n{ex.Message}" +
                        $"\n\nAny records processed so far have been saved to:\n{outputPath}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetRunning(false);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (!_running) return;
            _cts?.Cancel();
            AppendLog("Cancellation requested — finishing current ORS…");
            btnCancel.Enabled = false;
        }

        // ── Validation ────────────────────────────────────────────────────────

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
            if (string.IsNullOrWhiteSpace(txtSession.Text))
            {
                MessageBox.Show("Please enter a session name (e.g. A).", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (_credCache == null || _credCache.Count == 0)
            {
                MessageBox.Show("Please load credentials from the database first.",
                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            string locality = cmbLocality.SelectedItem?.ToString() ?? string.Empty;
            if (!_credCache.ContainsKey(locality))
            {
                MessageBox.Show(
                    $"No credential found for locality '{locality}' in the DroidCred table.\n" +
                    "Please reload credentials or select a different locality.",
                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private void SetRunning(bool running)
        {
            _running                    = running;
            btnStart.Enabled            = !running;
            btnCancel.Enabled           = running;
            btnBrowseInput.Enabled      = !running;
            btnBrowseOutput.Enabled     = !running;
            btnLoadCredentials.Enabled  = !running;
            cmbLocality.Enabled         = !running;
            grpConnection.Enabled       = !running; // disables child controls too
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
