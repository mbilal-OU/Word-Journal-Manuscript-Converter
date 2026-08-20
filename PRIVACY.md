# Privacy and research confidentiality

Word Journal Manuscript Converter is designed for unpublished scientific manuscripts.

## Manuscript processing

By default, manuscript processing is local.

- manuscript body text is not uploaded by the Python core
- citation mapping, journal analysis, template adaptation, report generation, and preservation verification run locally
- transformed documents are written to a new file
- the original manuscript is never overwritten
- failed transformations remove the unsafe output copy
- live EndNote, Zotero, and Mendeley fields are protected unless the user explicitly creates a separate static review copy

## Optional anonymous product analytics

Product analytics are optional.

The desktop app asks before enabling analytics. The Word add-in and website use their own consent controls. Declining analytics does not disable manuscript processing.

When enabled, analytics may include:

- feature or action name
- application version
- operating-system family/version
- anonymous installation identifier
- anonymous session identifier
- download button click
- page view
- approximate session duration

Analytics do **not** include:

- manuscript text
- manuscript filenames
- local file paths
- citation text
- bibliography text
- figures or tables
- manuscript hashes
- document metadata
- journal manuscript content

The desktop anonymous installation identifier is a randomly generated UUID stored in the user's local application settings. It is not derived from a device serial number, Microsoft account, email address, or manuscript.

Analytics are stored in a Supabase project controlled by the developer. Public clients can insert analytics events but cannot read analytics records.

## Feedback

Feedback is transmitted only when the user presses **Submit**.

Feedback may include:

- rating
- category
- feedback message
- optional contact email
- optional permission to be contacted
- application version and platform

Users should not paste confidential manuscript text into the feedback form.

## Update checks

The desktop app may contact GitHub Releases to check whether a newer product build is available. This request does not include manuscript content.

Automatic update checks can be disabled under **Privacy & analytics**.

## Word add-in

The Word task pane runs through Office.js and the hosted add-in page.

Citation and reference scanning happens inside the open Word document through Office.js. Citation jumps change the Word selection and do not send document text to the Word Journal Manuscript Converter analytics service.

If add-in analytics are enabled, only product action names such as `citation_scan` or `jump_reference` are sent.

## Website analytics

The project website asks before recording anonymous page views, download clicks, and session duration. No cookies are required for manuscript functionality.

## Security model

The public Supabase key embedded in the desktop application, website, and Word add-in is a publishable client key. Database Row Level Security allows public clients to insert approved analytics/feedback records but denies public reads, updates, and deletes.

No service-role or secret Supabase credential is embedded in public software.

## Future online features

Any future cloud processing mode must separately disclose:

- what manuscript content is uploaded
- where it is processed
- retention duration
- deletion behavior
- which third parties receive content
- whether content is used for model training

Cloud manuscript processing must not be silently enabled by a normal software update.
