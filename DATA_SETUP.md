# Contact + Visitor Data Setup

The GitHub Pages site is static, so the browser must not contain a GitHub token. This setup uses Google Apps Script as a small backend.

## What this adds

- Contact form continues to send the enquiry email to palakollu.santosh16@gmail.com.
- Every contact submission is also stored in a Google Spreadsheet.
- The spreadsheet has two tabs:
  - Contacts — name, mobile, email, subject, comments, timestamp and technical context.
  - Visitors — visit count data such as timestamp, page, language, referrer, screen size and browser user-agent.
- After each contact submission, the latest workbook is exported as data/wedding-website-data.xlsx and committed/updated in this GitHub repository.
- No visitor name is collected automatically. A person's identity is only known when they voluntarily submit the contact form.

## One-time Google Apps Script setup

1. Open Google Apps Script and create a new standalone project.
2. Copy the code from google-apps-script/Code.gs into the project.
3. In Project Settings → Script Properties, add:
   - GITHUB_TOKEN = a GitHub fine-grained personal access token with Contents: Read and write for Santosh93980/wedding-invitation.
   - GITHUB_REPO = Santosh93980/wedding-invitation
4. Deploy → New deployment → type Web app.
5. Execute as Me.
6. Who has access: Anyone.
7. Copy the Web App URL.
8. Open config.js in this repository and put that URL inside window.WEDDING_DATA_ENDPOINT = "..." .
9. Commit the config.js update.

## Important

Do not put the GitHub token into config.js, index.html, or contact.html.

The website analytics are intentionally limited to basic technical/visit information. It cannot reliably identify exactly who opened the wedding website unless they submit their details themselves.

The xlsx file is updated on each contact submission, so GitHub will contain the latest two-tab workbook.
