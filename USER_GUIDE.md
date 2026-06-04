# Office AI Assistant User Guide

This guide explains how to install and use the add-in in Excel, Word, and PowerPoint.

## 1. Install the add-in

Use this production manifest URL:

```text
https://633qaq.github.io/officezhushou/manifest.xml
```

### Excel on the web

1. Open Excel on the web.
2. Open or create a workbook.
3. Go to `Home` -> `Add-ins` -> `More Add-ins`.
4. Choose `Upload My Add-in`.
5. Upload `manifest.xml` from this project, or use the manifest URL if your Office tenant supports URL-based upload.
6. After installation, find `AI Assistant` on the `Home` ribbon.

### Excel desktop

The exact menu name can vary by Office version and tenant policy. The usual path is:

1. Open Excel.
2. Go to `Home` -> `Add-ins` -> `More Add-ins`.
3. Choose `Upload My Add-in` or `Shared Folder`, depending on your Office setup.
4. Select `manifest.xml`.
5. Open the `AI Assistant` button from the ribbon.

If your Office account is managed by an organization, the administrator may need to deploy the manifest through Microsoft 365 centralized deployment.

## 2. First-time setup

After the task pane opens:

1. Choose an AI provider.
2. Use `Direct Mode` if you want to enter your own AI API key in the task pane.
3. Use `Server Mode` only after a public backend has been deployed and configured.
4. Click `Save Settings`.
5. Click `Check Connection`.

At the moment, the published add-in is ready for direct mode. Server mode requires a separate HTTPS backend deployment.

## 3. Use it in Excel

The add-in currently works as a general Office AI task pane in Excel:

1. Type a topic or instruction.
2. Paste selected cell content into the context box if needed.
3. Choose an action such as polish, expand, translate, speaker notes, or design tips.
4. Copy the generated result back into the workbook.

Excel-specific range reading and writing is not implemented yet. That means the add-in opens in Excel, but it does not yet automatically read selected cells or write structured data back into a selected range.

## 4. Use it in Word or PowerPoint

1. Select text in the document or presentation.
2. Click `Read Selection`.
3. Choose an AI action.
4. Review the result.
5. Click `Write Back Office` or copy the result manually.

## 5. Common issues

### The add-in button does not appear

- Confirm the installed manifest is the latest production `manifest.xml`.
- Confirm the manifest includes Excel support.
- Restart Office after installing the add-in.
- If this is an organization account, ask the Microsoft 365 admin whether custom add-ins are allowed.

### The task pane opens but AI calls fail

- In direct mode, confirm the API key is correct.
- Some providers may block browser requests with CORS.
- In server mode, confirm a public HTTPS backend has been deployed and configured.

### Server login or document saving does not work

Server mode is not available until the backend is deployed. The GitHub Pages site only hosts the static task pane.

