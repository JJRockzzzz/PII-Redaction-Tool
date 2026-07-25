"""Reproducible labelled evaluation for the PII detection rules.

Run: python evaluate.py
"""

from __future__ import annotations

from redact import find_matches


CASES = [
    ("name", "Contact Person: Priya Kumar", {"person"}),
    ("email", "Send to priya.kumar@example.com", {"email"}),
    ("phone", "Telephone: +91 98765 43210", {"phone"}),
    ("address", "12 Example Avenue, Sample City, 123456, India", {"address"}),
    ("company", "Acme Private Limited", {"company"}),
    ("dob", "Date of Birth: January 1, 1990", {"dob"}),
    ("ip", "Client IP: 203.0.113.12", {"ip"}),
    ("ssn", "SSN: 123-45-6789", {"ssn"}),
    ("card", "Card: 4111 1111 1111 1111", {"card"}),
    ("cin", "U28129PN1979PLC141032", set()),
    ("registration", "SEBI registration INR000004058", set()),
    ("fiscal year", "Fiscal 2025-2026", set()),
    ("ordinary date", "Dated December 10, 2025", set()),
    ("order number", "Order 000013004", set()),
    ("invalid card", "4111 1111 1111 1112", set()),
    ("invalid ip", "192.168.300.1", set()),
    ("heading", "RED HERRING PROSPECTUS", set()),
    ("short identifier", "00135070", set()),
]


def main() -> None:
    tp = 0
    fp = 0
    fn = 0
    tn = 0
    failures: list[str] = []
    for label, text, expected in CASES:
        found = {kind for _, _, kind, _ in find_matches(text)}
        if expected:
            if expected <= found:
                tp += 1
            else:
                fn += 1
                failures.append(f"false negative: {label}; found {sorted(found)}")
            fp += len(found - expected)
        elif found:
            fp += 1
            failures.append(f"false positive: {label}; found {sorted(found)}")
        else:
            tn += 1
    accuracy = (tp + tn) / len(CASES)
    precision = tp / (tp + fp) if tp + fp else 0.0
    recall = tp / (tp + fn) if tp + fn else 0.0
    print(f"True Positives={tp} False Positives={fp} False Negatives={fn} True Negatives={tn}")
    print(f"Accuracy={accuracy:.1%} Precision={precision:.1%} Recall={recall:.1%}")
    if failures:
        raise SystemExit("\n".join(failures))


if __name__ == "__main__":
    main()
