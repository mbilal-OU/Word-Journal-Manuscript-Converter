# Word add-in

The Word Journal Manuscript Converter add-in provides live-safe citation/reference navigation inside Microsoft Word.

It is separate from the desktop converter. The desktop application handles journal retargeting, template adaptation, linked review-copy creation, and full DOCX preservation checks. The Word add-in is for navigation and lightweight inspection while a document is open.

## Why installation is more difficult in Early Access

The add-in is not yet published through Microsoft Marketplace. That means Microsoft treats the current installation as a **sideloaded test add-in** rather than a normal store installation.

This is why Early Access testers may need to upload a manifest in Word on the web or configure a trusted add-in catalog in desktop Word. The extra steps are a Microsoft testing/deployment requirement, not a requirement of the citation-navigation logic itself.

For stable public release, the intended user experience is:

1. Open Word.
2. Open **Add-ins**.
3. Find **Word Journal Manuscript Converter**.
4. Click **Add**.

That simpler flow requires Microsoft Marketplace submission and Microsoft review.

The desktop app should never silently weaken Word Trust Center settings just to make sideloading easier. Early Access setup can be guided, but security settings remain under the user's or organization's control.

## Early Access distribution

The add-in uses an Office add-in-only manifest:

`integrations/word-addin/manifest.xml`

The hosted installation page is:

`https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/word-addin/`

The task pane is hosted over HTTPS at:

`https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/addin/taskpane.html`

During Early Access, the add-in is intended to be **sideloaded for testing**, not treated as a production-store installation.

Microsoft's Office Add-ins deployment guidance distinguishes:

- sideloading for development/testing
- Microsoft Marketplace for public distribution
- Microsoft 365 integrated apps for organization-wide deployment

Official Microsoft documentation:

- Deploy and publish Office Add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish
- Sideload Office Add-ins for testing: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing
- Windows network-share sideloading: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
- Microsoft 365 admin-center deployment: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins
- Microsoft Marketplace publishing: https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource

## Easiest Early Access route: Word on the web

For individual testing, the simplest path is usually Word on the web:

1. Open the manuscript in Word on the web.
2. Open **Add-ins**.
3. Choose the option to upload or sideload a custom add-in.
4. Select the supplied `manifest.xml`.
5. Open Word Journal Manuscript Converter from the add-ins menu.

The exact Microsoft 365 labels can vary by build. Follow Microsoft's current sideloading page above if the labels shown in Word differ.

## Desktop Word testing on Windows

Microsoft documents a trusted network-share catalog method for testing task-pane add-ins on Windows.

This is a development/testing method, not the intended production-distribution method.

Use the manifest supplied in the release package and follow Microsoft's network-share sideloading instructions. The desktop application also exposes a guided setup window so testers can locate the local manifest and open the current installation guide without searching through the package manually.

## Organization deployment

A Microsoft 365 administrator can deploy a custom Office Add-in to selected users or groups through **Settings > Integrated apps** in the Microsoft 365 admin center.

This is appropriate for an internal pilot if the institution allows custom add-in deployment.

## Stable public distribution

For the stable public product, the intended route is **Microsoft Marketplace**.

That provides:

- normal in-product discovery
- a standard installation experience
- centrally managed add-in updates
- Microsoft certification/validation

Marketplace publication requires a Partner Center account and Microsoft review, so it remains a launch-readiness step rather than something the repository can complete automatically.

## Privacy

Citation and reference scans run against the open Word document through Office.js.

The add-in does not send manuscript text to Word Journal Manuscript Converter analytics.

Optional add-in analytics can record action names such as:

- add-in opened
- citation scan
- integrity check
- jump to citation
- jump to reference

The analytics toggle is off until the user enables it.

## Independence

Word Journal Manuscript Converter is independently developed and is not affiliated with or endorsed by Microsoft, Clarivate/EndNote, Zotero, Mendeley, or any journal publisher.
