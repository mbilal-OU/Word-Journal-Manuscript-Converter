# Word add-in

The Word Journal Manuscript Converter add-in provides live-safe citation and reference navigation inside Microsoft Word.

It is separate from the desktop converter. The desktop application handles journal conversion, template adaptation, linked review-copy creation, and DOCX preservation checks. The Word add-in is for navigation and lightweight inspection while a document is open.

## Why Early Access installation is more technical

The add-in is not yet published through Microsoft Marketplace. Microsoft therefore treats the current build as a sideloaded test add-in rather than a normal store installation.

This affects all Early Access testers. It is not specific to one computer, citation manager, or institution.

The stable distribution target is the normal Word experience:

1. Open Word.
2. Open Add-ins.
3. Search for Word Journal Manuscript Converter.
4. Select Add.

That path requires Microsoft Marketplace submission and review and is not yet available.

## Fastest Early Access route

For an individual tester, the simplest route is Word on the web:

1. Open Word on the web.
2. Open Add-ins.
3. Use the custom or upload add-in option available in your Microsoft 365 build.
4. Upload the supplied `manifest.xml`.
5. Open a manuscript and launch Citation Navigator.

The exact labels can vary across Microsoft 365 builds. Use Microsoft's current sideloading guidance if your interface uses different wording.

The manifest is included in release packages at:

`word-addin/manifest.xml`

Repository source:

`integrations/word-addin/manifest.xml`

Hosted task pane:

`https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/addin/taskpane.html`

Hosted setup page:

`https://mbilal-ou.github.io/Word-Journal-Manuscript-Converter/word-addin/`

## Windows desktop testing

Microsoft also documents a trusted catalog method for Windows desktop Word. This is a developer and testing workflow and is more technical than Marketplace installation.

Word Journal Manuscript Converter does not silently change Word Trust Center settings. On university or company Microsoft 365 accounts, organization policy may also restrict custom add-ins.

## Organization deployment

A Microsoft 365 administrator can deploy a custom Office add-in to selected users or groups through the Microsoft 365 admin center when institutional policy permits it.

## Microsoft documentation

- Deploy and publish Office Add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish
- Sideload Office Add-ins for testing: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing
- Windows network-share catalog: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
- Microsoft 365 integrated apps: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins
- Microsoft Marketplace publishing: https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource

## Privacy

Citation and reference scans run against the open Word document through Office.js.

The add-in does not send manuscript text to Word Journal Manuscript Converter analytics.

Optional add-in analytics may record product events such as add-in opened, citation scan, integrity check, jump to citation, or jump to reference. The analytics toggle is off until the user enables it.

## Independence

Word Journal Manuscript Converter is independently developed and is not affiliated with or endorsed by Microsoft, Clarivate/EndNote, Zotero, Mendeley, or any journal publisher.
