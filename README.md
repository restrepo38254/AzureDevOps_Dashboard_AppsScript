# Azure DevOps Pipeline Dashboard (Apps Script)

This project contains a Google Apps Script that shows a dashboard for Azure DevOps pipelines. The script fetches pipeline data using the Azure DevOps REST API and renders metrics in a web app.

## Installation

1. **Clone** this repository and install the development dependencies:
   ```bash
   npm install
   ```
2. **Create** a new Apps Script project (or open an existing one) and copy the files from the `src` directory and `appsscript.json` into your script. Tools such as [`clasp`](https://github.com/google/clasp) can be used to push the files.

## Configuration

Before deploying you need to set the authentication token and optionally adjust some limits.

1. **Set the PAT token**

   In the Apps Script editor open the script console and run the following function once:
   ```javascript
   setPatToken('YOUR_PAT_TOKEN');
   ```
   This stores the token in the script properties so that subsequent API calls can authenticate.

2. **Optional limits**

   The dashboard limits the number of pipelines and runs that it loads. Edit `CONFIG.maxPipelines` and `CONFIG.maxRuns` at the top of `src/Code.ts` to suit your project size.

## Running the Dashboard

Deploy the script as a web app from the Apps Script editor ("Deploy" → "New deployment"). Open the deployment URL to view the dashboard. The preview can also be launched locally with:

```bash
npm run preview
```

This starts a local server that serves `src/Index.html` for quick interface previews.

### Known limitations

* The dashboard only loads up to the number of pipelines and runs configured in `CONFIG`.
* If many pipelines are fetched the script may approach Apps Script execution time limits.
* Ensure the PAT token has permissions to read pipeline information in your Azure DevOps organization.

