# Word add-in

The Word Journal Manuscript Converter add-in provides live-safe citation/reference navigation inside Microsoft Word.

It is separate from the desktop converter. The desktop application handles journal retargeting, template adaptation, linked review-copy creation, and full DOCX preservation checks. The Word add-in is for navigation and lightweight inspection while a document is open.

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

## Option A: Word on the web

For individual Early Access testing, use Microsoft's sideloading workflow for Word on the web and upload the provided manifest when prompted.

The exact Office UI can vary by Microsoft 365 build. Follow Microsoft's sideloading page above if the labels shown in Word differ.

## Option B: Windows desktop testing

Microsoft documents a trusted network-share catalog method for testing task-pane add-ins on Windows.

This is a development/testing method, not the intended production-distribution method.

Use the manifest supplied in the release package and follow Microsoft's network-share sideloading instructions.

## Option C: organization deployment

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
