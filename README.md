# PII Redaction Tool

`redact.py` script redacts PII from a DOCX file and writes a new DOCX with the specified requirements, leaving the actual source unmodified. 

Usage:

python redact.py "Red Herring Prospectus.docx" "Red Herring Prospectus - Redacted.docx"

python evaluate.py

The implementation is a conservative hybrid of regular expressions (regex) and usual document-context rules. It handles all e-mail addresses, Indian telephone numbers, IPv4 addresses, SSNs, Credit Card no.s, physical addresses, organisation names, titled or contact person names, and dates only when the paragraph explicitly identifies them as a date of birth. Replacements are deterministic within each run, so repeated values receive the same alternative determined earlier.

The script parses the body text, nested tables, headers, and footers. It replaces character spans across Word runs so most formatting is expected to be retained. At places where a PII value crosses differently formatted runs, the replacement takes the formatting of its first character only. It does not inspect any images, scanned pages, text boxes, or encrypted content present inside the document. All names, companies, and addresses are inherently ambiguous, the deliberately strict heuristics trade some recall score for lower false-positive risk. For production, I would add a reviewed entity dictionary and an NER model, then route uncertain detections to human review for better accuracy.

`evaluate.py` is a reproducible labelled test harness that was used to determine the accuracy, precision, and recall figures documented within `evaluation_report.md`.
