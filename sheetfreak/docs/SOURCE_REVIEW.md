# Source review and provenance

The source in this folder was copied from the owner-supplied SheetFreak_v01 ZIP. Version identifiers inside the archive are 0.2.1. The ZIP container name is not the version identifier.

Included unchanged: the two production VBA modules, ribbon XML, two test modules, manual checklist, recorded test logs, build helper, changelog and MIT license. The README was rewritten for a source-only public showcase. No prior-job macros, reference models, compiled add-in, binary checksum file or source archive is included.

The archive was inspected without loading the add-in or running macros, tests or its build script. No credentials, client data, personal paths or network calls were found in the included text source. Historical test results are included as supplied and were not independently rerun.

The excluded XLAM contains a local absolute path in workbook XML even though its creator and last-editor fields have been generalized. Its compiled VBA project was not decompiled or matched to the text modules. Do not add that binary to a public repository without a fresh build and metadata review.
