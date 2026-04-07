# InboxCraft Skills Test Suite

This folder contains standalone PowerShell scripts corresponding to the code templates defined in each skill's `SKILL.md` file. It allows developers and testers to rapidly verify the underlying Outlook COM interactions natively without having to invoke the LLM agent to generate the scripts.

## Usage

To quickly test the non-destructive skills on your local Outlook instance, open PowerShell and run:

```powershell
.\run_all_tests.ps1
```

This master runner will sequentially execute the safe, read-only testing scripts located in this directory:
- `test_show_rules.ps1`
- `test_show_folders.ps1`
- `test_find_large_folders.ps1`
- `test_export_categories.ps1`

## Destructive / Interactive Tests

Several scripts require user input or perform destructive actions. The `run_all_tests.ps1` script explicitly **does not** run these automatically. 

You must run these manually to test them:
- **`test_export_rules.ps1`**: Exports your rules to a CSV file.
- **`test_export_folders.ps1`**: Exports your folder tree to a CSV file.
- **`test_clean_empty_folders.ps1`**: Recursively deletes empty folders. *(Note: This script has `$DryRun = $true` set by default to prevent accidental deletion during testing).*
- **`test_disable_all_rules.ps1`**: Unchecks/pauses all active Inbox rules. *(Note: This script has the save functionality commented out by default).*

## Schema Validation

To validate the repository's skills against the Agent Skills specification, use the following npm command from the repository root:

```bash
npm run validate
```
This runs the `scripts/validate.js` utility.

## Contributing

If you add a new skill to the `skills/` directory, please remember to extract its PowerShell template and add a corresponding `test_my_new_skill.ps1` file to this folder so it can be verified easily.
