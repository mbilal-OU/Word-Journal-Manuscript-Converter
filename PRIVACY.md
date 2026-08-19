# Privacy and research confidentiality

Word Journal Manuscript Converter is designed for unpublished scientific manuscripts.

## Default policy

- manuscript processing is local;
- no manuscript body is transmitted by the Python core;
- inspection, citation mapping, retargeting, and preservation verification require no network connection;
- automatic transformations write a new output file rather than overwriting the source;
- failed transformations remove the unsafe output copy;
- logs and reports should contain counts, paths, hashes, rule outcomes, and short structural labels, not full manuscript prose by default;
- no manuscript-content telemetry is part of the current core.

## Word add-in

The current Word task-pane starter performs a lightweight check inside Word and does not send manuscript text to a Word Journal Manuscript Converter service. Office.js itself is loaded from Microsoft's hosted Office library as required by the add-in development model.

## Future online features

Fetching public journal guidelines, DOI metadata, or citation metadata should send the minimum public identifier required. The manuscript body should not be sent merely to retrieve journal rules.

Any future cloud processing mode must separately disclose:

- what is uploaded;
- where it is processed;
- retention duration;
- whether third-party APIs receive content;
- deletion behavior;
- whether content is used for model training.
