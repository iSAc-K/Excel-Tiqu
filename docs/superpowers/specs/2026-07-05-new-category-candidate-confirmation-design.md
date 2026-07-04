# New Category Candidate Confirmation Design

## Context

The extractor currently classifies orders by matching configured category keywords against recognition names. Recognition names already include the order folder name before archive, date-folder, and Excel file names. If no keyword matches, the file is classified as `未分类`; those rows are excluded from the formal output and the source Excel is copied to `未分类Excel`.

This behavior is safe, but it creates extra manual work when a folder name clearly contains a new product category that is not yet present in `category_config.json`, such as `方白名片架`.

## Goal

Add a semi-automatic workflow that discovers possible new categories from unclassified folder names, shows them in the GUI for confirmation, and saves user-approved category and prefix rules for future runs.

The tool must not silently add new categories to the formal configuration without user confirmation.

## Non-Goals

- Do not classify unconfirmed candidates into the formal output during the same unattended extraction run.
- Do not use fuzzy matching, AI recognition, or external services.
- Do not change existing keyword priority rules for already configured categories.
- Do not remove the current `未分类Excel` safety behavior.

## Candidate Extraction Rules

When keyword matching returns `未分类`, the extractor should try to produce a candidate from the same ordered recognition names it already uses:

1. order folder name
2. archive name or date directory name
3. Excel file name

For each name, the candidate parser should remove structural segments and keep the most likely product category text:

- leading sequence ranges, such as `18~19`, `23~24`, `1~6`
- date segments, such as `0620`, `0625`, `6.25`, `06-25`
- quantity segments, such as `20单`, `3个`, `29单`, `13个`
- empty separators and repeated punctuation

Examples:

| Raw name | Candidate before prefix cleanup |
| --- | --- |
| `18~19-0620-0625-方黑名片架-20单-3个` | `方黑名片架` |
| `17-0625-WZY-纯木名片架-5单5个` | `WZY-纯木名片架` |
| `14-0625-HAL小钢片-4单5个` | `HAL小钢片` |
| `33-6.25-CSJ-MA88钥匙扣-8单-23个` | `CSJ-MA88钥匙扣` |

If the parser cannot find a meaningful text segment after removing structure, it should not create a candidate.

## Prefix Rules

Add a configurable supplier/channel prefix list. The prefix list is separate from category keywords because it describes noise to remove, not categories to classify.

Prefix cleanup should support both common forms:

- separated prefix: `WZY-纯木名片架` -> prefix `WZY`, category `纯木名片架`
- attached prefix: `HAL小钢片` -> prefix `HAL`, category `小钢片`

The GUI must allow users to add or edit prefixes when a candidate is not cleaned correctly. Saved prefixes apply to future candidate extraction.

Seed the prefix list as empty for existing users. Users add prefixes from the candidate review UI as they encounter unknown supplier/channel text.

## GUI Workflow

Add a new GUI entry point for reviewing discovered category candidates. The entry point can be a button or panel near the existing category keyword configuration.

The candidate review list should show:

- original source name
- detected prefix, if any
- proposed category
- source Excel or folder path for traceability
- action status

Each row should support:

- save as new category
- edit prefix
- edit proposed category
- ignore

If a prefix is unknown, the user can enter the prefix portion. The GUI should immediately show the resulting category after removing that prefix.

Saving a candidate should update persistent configuration with:

- a prefix rule, when the user entered or confirmed a prefix
- a new category with the category itself as its first keyword

Move `category_config.json` to a backward-compatible structured format when the first prefix or candidate save happens. Old files that use the current top-level category map must still load normally. When saving after this feature is used, write:

```json
{
  "prefixes": ["HAL"],
  "categories": {
    "小钢片": ["小钢片"]
  }
}
```

For example, confirming `HAL小钢片` with prefix `HAL` writes `HAL` to `prefixes` and `小钢片: ["小钢片"]` under `categories`. Existing top-level category maps should be interpreted as `{"prefixes": [], "categories": <old map>}` during load.

## Processing Flow

1. Extraction runs with current keyword classification.
2. If a file is still `未分类`, candidate extraction runs against recognition names.
3. Candidate metadata is stored in the run result and processing report.
4. Unconfirmed candidates remain excluded from the formal output, matching current `未分类` behavior.
5. The GUI shows candidate rows for user confirmation.
6. After the user saves candidates, the updated config is used by the next extraction run.

The first implementation can require rerunning extraction after saving candidates. A later enhancement may add a dedicated "重新识别" action, but that is not required for this design.

## Error Handling

- If configuration loading fails, keep the existing fallback behavior and disable saving candidate changes until the config is readable or reset.
- If a proposed category already exists, add the candidate text as a keyword only if it is not already present.
- If a prefix already exists, do not duplicate it.
- If the user clears the proposed category, block saving and show a validation message.
- If two candidates produce the same category, save one category and merge keywords without duplicates.

## Reporting

The processing report should include a candidate section or columns that make unclassified candidates visible after a run:

- original source name
- proposed category
- detected prefix
- whether it was already confirmed
- reason it still stayed unclassified

This helps users review candidates even if they do not open the GUI immediately.

## Testing

Add focused tests for:

- extracting `方黑名片架` from dated and quantity-suffixed folder names
- cleaning separated prefixes such as `WZY-纯木名片架`
- cleaning attached prefixes such as `HAL小钢片`
- not producing candidates from names that only contain dates, counts, or order numbers
- preserving existing keyword matching priority
- keeping unconfirmed candidates out of the formal output
- saving confirmed category and prefix rules without duplicating existing config entries
- loading old `category_config.json` files without prefixes

GUI behavior should be covered by unit-level config and candidate model tests where possible, plus a manual smoke test for the review dialog.
