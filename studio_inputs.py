"""Validation and filtering for the reconciliation workspace."""
import re

import pandas as pd

ISSUE_TYPES = ("All exceptions", "Missing fields", "Possible repeats", "Unresolved identity")


def parse_domains(value: str) -> tuple[list[str], list[str]]:
    """Accept pasted domain lists; reject URLs and email addresses explicitly."""
    domains, invalid = [], []
    label = re.compile(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?")
    for entry in re.split(r"[\n,;]+", value):
        candidate = entry.strip().removeprefix("@").lower()
        if not candidate:
            if entry.strip():
                invalid.append(entry.strip())
            continue
        try:
            domain = candidate.encode("idna").decode("ascii")
        except UnicodeError:
            invalid.append(entry.strip())
            continue
        if len(domain) > 253 or not all(label.fullmatch(part) for part in domain.split(".")):
            invalid.append(entry.strip())
        elif domain not in domains:
            domains.append(domain)
    return domains, invalid


def parse_markers(value: str) -> list[str]:
    return list(dict.fromkeys(item.strip().lower() for item in value.splitlines() if item.strip()))


def filter_exceptions(records: pd.DataFrame, query: str, issue: str) -> pd.DataFrame:
    result = records
    if issue == "Missing fields":
        result = result.loc[result["Missing Critical Fields"].fillna("").astype(str).str.strip().ne("")]
    elif issue == "Possible repeats":
        result = result.loc[result["Exact Repeat Review"].fillna("").astype(str).str.strip().ne("")]
    elif issue == "Unresolved identity":
        result = result.loc[result["Participant Key"].fillna("").astype(str).str.startswith("UNRESOLVED:")]
    if query.strip():
        mask = result.fillna("").astype(str).apply(
            lambda column: column.str.contains(query.strip(), case=False, regex=False)
        ).any(axis=1)
        result = result.loc[mask]
    return result
