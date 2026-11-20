# Changelog

## 1.2.1

Fixes from a code review.

- Excel export keeps emoji and other characters above U+FFFF; they were silently dropped.
- Allowed Values in exports match on table and column, so a view showing `account.statuscode` no longer brings in the choices of `contact.statuscode`.
- A malformed form or view definition is reported as an error instead of giving an empty or shortened list.
- Long Dataverse result lists are read page by page instead of being cut off at the first page.
- The "different environment" warning can no longer be suppressed by a look-alike host such as `make.powerapps.com.example.org`.
- Columns of type Table Name show their choices; multi-select choice columns show their default; labels missing in the chosen language fall back to English.
- CSV and copied text can no longer start a spreadsheet formula (text starting with `=`, `+`, `-` or `@` gets a leading apostrophe). This also stops negative ranges such as `-2,147,483,648 to 2,147,483,647` from showing as errors.
- Markdown export escapes `<`, so text from Dataverse cannot turn into HTML.
- Two forms given the same name are kept apart as "Name (2)", with a note under the name box.
- Columns the tool shows by itself (Visible, Notes, Table, Also on) are hidden again when no longer relevant and are not remembered; columns you tick yourself are.
- Forms of the same table share the column metadata already read, so adding more forms is faster.
- Smaller fixes: one export at a time, focus stays on the page when the form chooser opens or closes or a form is removed, stored settings are checked before use.
- Development: tests are type-checked and every test file runs; CI runs build, tests and validation on each push and pull request.

## 1.2.0

- **Add form** replaces the Build dictionary button and can be used repeatedly: collect several forms and views in one dictionary. Exports cover all of them, with one Excel worksheet per form.
- **Form / Excel tab name** box: name each form; the name is used on the form strip and as the worksheet name.
- **Also on** column: marks fields that also appear on other forms in the dictionary.
- A failed or cancelled read no longer clears the dictionary on screen.
- A × in the address box clears it.
- Description reworded in the style of Dataverse M Builder.

## 1.1.0

- Tab buttons are multi-select: show any combination of tabs; exports follow the selection.
- No more false "different environment" warning when the connection uses the Web API host (`org.api.crm4.dynamics.com`).
- Manifest: contributors in object form, minimum ToolBox version 1.2.7.
- README: screenshot, requirements and AI assistance statement.

## 1.0.1

- New icon in the style of Dataverse M Builder.

## 1.0.0

- First release.
