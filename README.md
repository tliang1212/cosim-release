# Cosim Release Dashboard

This repository contains a Databricks App built with Streamlit. It is designed
to help customers filter a cosim release plan by cadence, program, and domain,
then review release timing, planned content, and known issues.

The app runs locally with:

```bash
streamlit run app.py
```

It can then be deployed to Databricks Apps with the same entry point through
`app.yaml`.

## Files

| File | Purpose |
| --- | --- |
| `app.py` | Streamlit dashboard application |
| `requirements.txt` | Python packages installed locally and by Databricks Apps |
| `app.yaml` | Databricks Apps runtime configuration |
| `databricks.yml` | Optional Databricks Asset Bundle configuration |

## Expected Excel columns

The app works best when the Excel sheet has columns similar to these:

- `Cadence`
- `Program`
- `Domain`
- `Release` or `Milestone`
- `Release Date`, `Target Date`, or `Timing`
- `Content`
- `Issues` or `Known Issues`
- `Owner`
- `Status`
- `Notes`

If your Excel headers are different, use the **Column mapping** section in the
app sidebar to map your headers to the dashboard fields. This is useful when the
current Excel layout has merged business terms or different naming conventions.

## Step 1: Prepare the Excel file

For quick local testing, upload the Excel file in the app sidebar.

For a shared Databricks App, place the Excel file somewhere the app can read it.
Recommended options:

1. Unity Catalog Volume, for example:

   ```text
   /Volumes/<catalog>/<schema>/<volume>/cosim_release_plan.xlsx
   ```

2. DBFS path, for example:

   ```text
   dbfs:/FileStore/cosim/cosim_release_plan.xlsx
   ```

3. A local project file for prototypes:

   ```text
   data/cosim_release_plan.xlsx
   ```

For production use, prefer a Unity Catalog Volume and grant the Databricks App
service principal permission to read the file.

## Step 2: Develop locally in VS Code

From the repository root:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

On Windows, use `py -m venv .venv` and activate with
`.venv\Scripts\activate`.

Open the local Streamlit URL that appears in the terminal. Upload your Excel
file, select the correct worksheet, and adjust the column mapping if needed.

## Step 3: Point the app at the Databricks Excel file

Edit `app.yaml` and update `COSIM_EXCEL_PATH`.

Example for a Unity Catalog Volume:

```yaml
env:
  - name: 'COSIM_EXCEL_PATH'
    value: '/Volumes/main/release_planning/files/cosim_release_plan.xlsx'
```

Example for DBFS:

```yaml
env:
  - name: 'COSIM_EXCEL_PATH'
    value: 'dbfs:/FileStore/cosim/cosim_release_plan.xlsx'
```

The app also supports sidebar upload, but that is intended for testing. A shared
dashboard should use a managed Databricks file path so every user sees the same
release plan.

## Step 4: Configure Databricks CLI

Install or update the Databricks CLI, then authenticate to the GM Databricks
workspace:

```bash
databricks -v
databricks auth login --host https://<your-workspace-url>
```

If you use profiles:

```bash
databricks configure --host https://<your-workspace-url> --profile gm
```

Then pass `--profile gm` on CLI commands.

## Step 5: Create the Databricks App

In the Databricks UI:

1. Open the Databricks workspace.
2. Open **Databricks Apps**.
3. Create a new app named `cosim-release-dashboard`.
4. Confirm the app service principal has access to the Excel file location.

If your CLI supports app creation, you can also create it from the terminal:

```bash
databricks apps create cosim-release-dashboard
```

## Step 6: Deploy from VS Code

From the repository root:

```bash
databricks apps deploy cosim-release-dashboard --source-code-path .
```

If you configured a profile:

```bash
databricks apps deploy cosim-release-dashboard --source-code-path . --profile gm
```

Databricks Apps will install `requirements.txt`, read `app.yaml`, and start the
dashboard with:

```bash
streamlit run app.py
```

## Optional: Deploy with Databricks Asset Bundles

Edit `databricks.yml` and replace:

```yaml
host: https://replace-me.cloud.databricks.com
```

with your GM workspace URL. Then validate and deploy the bundle resources:

```bash
databricks bundle validate
databricks bundle deploy
databricks apps deploy cosim-release-dashboard --source-code-path .
```

Use `--target prod` when you are ready to deploy the production target.

## Dashboard behavior

The app provides:

- Cadence, program, and domain filters
- Optional upcoming-only filter
- Text search across timing, content, issues, owner, status, and notes
- Release timing chart
- Filtered release plan table with CSV download
- Domain detail view
- Known issues view

## Troubleshooting

- **Workbook not found**: verify `COSIM_EXCEL_PATH` in `app.yaml` and confirm
  the Databricks App service principal can read that path.
- **No filters show up**: verify the selected worksheet and column mapping.
- **Dates do not chart correctly**: make sure the mapped timing column contains
  actual Excel dates or parseable date strings.
- **Deployment fails installing packages**: redeploy after confirming
  `requirements.txt` is present at the repository root.
- **App starts but users cannot see data**: grant file or volume permissions to
  the app's service principal, not only your personal user account.
