

Snippet 6 — MainForm.cs — Session B credentials + wire MhiSession
Step 6a — Add two new TextBox fields with the other field declarations at the top of the class:
csharp// In MainForm.cs — add alongside other private field declarations (after txtSession fields)

private Label   lblUnetId;
private TextBox txtUnetId;
private Label   lblUnetPass;
private TextBox txtUnetPass;
Step 6b — In InitializeComponent, after the txtSession block inside the GroupBox section, add:
csharp// In InitializeComponent(), after txtSession config, before lblLocality

lblUnetId       = new Label();
txtUnetId       = new TextBox();
lblUnetPass     = new Label();
txtUnetPass     = new TextBox();
And in the properties section (after txtSession.Text = "A";):
csharp// Session B Unet credentials
lblUnetId.Text     = "Unet ID (Sess B):";
lblUnetId.Left     = 10;
lblUnetId.Top      = 46;        // below session name row
lblUnetId.Width    = 120;
lblUnetId.AutoSize = false;

txtUnetId.Left  = 135;
txtUnetId.Top   = 44;
txtUnetId.Width = 120;

lblUnetPass.Text     = "Unet Pass:";
lblUnetPass.Left     = 270;
lblUnetPass.Top      = 46;
lblUnetPass.Width    = 80;
lblUnetPass.AutoSize = false;

txtUnetPass.Left                  = 355;
txtUnetPass.Top                   = 44;
txtUnetPass.Width                 = 120;
txtUnetPass.UseSystemPasswordChar = true;
Shift lblCredStatus.Top down by 24 to avoid overlap:
csharplblCredStatus.Top = 76;   // was 52
And add the new controls to grpConnection.Controls.Add(...):
csharpgrpConnection.Controls.Add(lblUnetId);
grpConnection.Controls.Add(txtUnetId);
grpConnection.Controls.Add(lblUnetPass);
grpConnection.Controls.Add(txtUnetPass);



Step 6c — In ValidateInputs(), add before the return true:
csharp// In ValidateInputs(), add before return true

if (string.IsNullOrWhiteSpace(txtUnetId.Text) || string.IsNullOrWhiteSpace(txtUnetPass.Text))
{
    MessageBox.Show("Please enter Unet ID and Password for Session B (MHI lookup).",
        "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    return false;
}




Step 6d — In BtnStart_Click, replace the Task.Run lambda body:
csharp// In BtnStart_Click Task.Run lambda — replace with:

string unetId   = txtUnetId.Text.Trim();
string unetPass = txtUnetPass.Text;

results = await Task.Run(() =>
{
    // Session A — Payloc / ORS navigation
    var session = new AS400Session(sessionName)
    {
        ActionDelayMs     = actionDelay,
        PageScrollDelayMs = scrollDelay
    };
    session.SetCredential(cred, locality);
    session.Connect();

    // Session B — Seamless login for MHI (created once, reused for all ORS IDs)
    var mhi = new MhiSession(unetId, unetPass);
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