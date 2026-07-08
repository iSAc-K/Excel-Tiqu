# Save Candidate To Existing Category Design

## Context

The current new-category candidate window supports two outcomes: save the candidate as a new category, or ignore it. Some candidates are not true new products. They are unknown keywords or supplier-prefixed names that should be attached to an existing category.

Example:

- Candidate source: `7.5-CBZ-心形钥匙扣-1单-1个`
- Extracted prefix: `CBZ`
- Editable candidate keyword: `心形钥匙扣扣` or a user-corrected value
- Existing category target: `心形钥匙扣`

The user needs an explicit "save to existing category" path so these candidates do not create noisy new categories.

## Design

Add an "已有品类" selector to `NewCategoryCandidatesWindow`, populated from `category_config.json` categories.

Add a button:

- `保存到已有品类`

Keep the existing buttons:

- `按前缀重新切分`
- `保存为新品类`
- `忽略此候选`

The editable "品类" field keeps its current role as the candidate keyword/name field. It is not automatically normalized beyond prefix splitting. If a candidate has an extra character such as `心形钥匙扣扣`, the user can correct it in the field before saving.

## Data Flow

When the user clicks `保存到已有品类`:

1. Read the selected existing category from the selector.
2. Read `prefix` from the prefix field and preserve it in `prefixes` if non-empty.
3. Read the editable candidate keyword from the current category field.
4. Append that keyword to the selected existing category's keyword list if it is not already present.
5. Save through the structured category config model so `prefixes` are preserved.
6. Update the live candidate status to `已保存到已有品类：<品类>`.

When the user clicks `保存为新品类`, keep the current behavior: create or update a category named by the editable category field and add the same keyword to it.

## Error Handling

- Disable or reject `保存到已有品类` if there are no configured categories.
- Show a warning if no existing category is selected.
- Show a warning if the candidate keyword is empty.
- Preserve the existing invalid-config protection: do not overwrite corrupt `category_config.json`.

## Testing

Add regression coverage for:

- Saving a candidate keyword to an existing category preserves existing keywords.
- Saving to an existing category does not create a new category named after the candidate.
- Saving to an existing category still records the prefix in structured config.
- GUI-level helper behavior mutates the live candidate status so reopening the window does not show stale `待确认`.
