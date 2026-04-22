# EntityKeywordControl

A **ReactJS / Fluent UI PCF (Power Apps Component Framework)** control that presents a **Many-to-Many related list of keyword records** as a rich, interactive tag group — similar to a multi-select but with full support for per-record metadata (colour, icon, grouping, description, and sort order).

---

## Features

| Feature | Description |
|---|---|
| **Tag display** | Associated keywords rendered as dismissible Fluent UI v9 `Tag` components |
| **Colour coding** | Each keyword record can store a hex colour used to tint its tag |
| **Icon support** | Keywords can specify an image URL or Fluent icon name displayed on the tag |
| **Grouping** | Keywords can be grouped by a category field inside the picker dialog |
| **Description tooltips** | Hover a tag to see the keyword's description field |
| **Sort order** | Optional integer field controls display order within the tag group |
| **Overflow badge** | Configurable `maxVisibleTags` clips long lists with a `+N more` badge |
| **Search & filter** | Full-text search across label, group and description in the picker dialog |
| **Associate** | Check-box multi-select in the picker; all selected records are bulk-associated |
| **Disassociate** | Dismiss button (×) on each tag removes the M2M association |
| **Create (optional)** | When `allowCreate` is enabled, a "Create" option appears for unmatched searches |
| **Read-only mode** | Respects the form's disabled/read-only state and hides action buttons |
| **Accessibility** | Full keyboard navigation, ARIA labels, and screen-reader announcements |

---

## Architecture

```
EntityKeywordControl/
├── ControlManifest.Input.xml      PCF manifest – properties & resources
├── index.ts                       PCF ReactControl entry point
├── components/
│   ├── types.ts                   Shared TypeScript interfaces
│   ├── EntityKeywordControl.tsx   Root React component (state, WebAPI, layout)
│   ├── KeywordTag.tsx             Single tag + overflow badge components
│   └── KeywordPickerPanel.tsx     Modal dialog for searching & associating keywords
├── __tests__/
│   └── EntityKeywordControl.test.tsx  Jest/RTL unit tests
├── package.json
├── tsconfig.json
├── babel.config.json
└── EntityKeywordControl.pcfproj  MSBuild project file
```

---

## Manifest Properties

| Property | Type | Required | Description |
|---|---|---|---|
| `relationshipName` | `SingleLine.Text` | ✅ | Schema name of the M2M relationship, e.g. `xrm_account_keyword` |
| `relatedEntityName` | `SingleLine.Text` | ✅ | Logical name of the keyword entity, e.g. `xrm_keyword` |
| `relatedEntitySetName` | `SingleLine.Text` | ✅ | OData entity-set name (plural), e.g. `xrm_keywords` |
| `navigationPropertyName` | `SingleLine.Text` | ✅ | Navigation property on the host entity for the M2M, e.g. `xrm_account_keyword_xrm_keyword` |
| `labelFieldName` | `SingleLine.Text` | ✅ | Field on the keyword entity used as the visible tag label, e.g. `xrm_name` |
| `colorFieldName` | `SingleLine.Text` | — | Field containing a hex colour string (`#E8A202`) to tint the tag |
| `iconFieldName` | `SingleLine.Text` | — | Field containing a Fluent icon name or absolute image URL |
| `groupFieldName` | `SingleLine.Text` | — | Field used to group keywords in the picker dialog |
| `descriptionFieldName` | `SingleLine.Text` | — | Field shown as a tooltip on the tag and subtitle in the picker |
| `sortOrderFieldName` | `SingleLine.Text` | — | Integer field controlling tag display order (ascending) |
| `allowAssociate` | `TwoOptions` | — | Show the "Add keyword" button (default: `true`) |
| `allowDisassociate` | `TwoOptions` | — | Show dismiss (×) buttons on tags (default: `true`) |
| `allowCreate` | `TwoOptions` | — | Show a "Create" option in the picker for unmatched searches (default: `false`) |
| `searchPlaceholder` | `SingleLine.Text` | — | Placeholder text in the picker search input |
| `maxVisibleTags` | `Whole.None` | — | Max tags shown inline before `+N more`; `0` = show all (default: `0`) |

---

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) ≥ 16
- [Power Apps CLI (`pac`)](https://docs.microsoft.com/en-us/powerapps/developer/data-platform/powerapps-cli) for building and deploying

### Install dependencies

```bash
npm install
```

### Run tests

```bash
npm test
```

### Start local test harness

```bash
npm start
```

This launches the PCF test harness at `http://localhost:8181` where you can exercise the control with mock data.

### Build

```bash
npm run build
```

Output is placed in `out/controls/EntityKeywordControl/`.

### Package and deploy

```bash
# Create solution zip
pac solution init --publisher-name XRMatic --publisher-prefix xrm
pac solution add-reference --path .
msbuild /t:build /restore

# Import to Dataverse
pac solution import --path bin/Release/*.zip
```

---

## Dataverse Setup

### 1 – Create the keyword entity

Create a custom table (e.g. `xrm_keyword`) with at minimum:

| Column | Type | Notes |
|---|---|---|
| `xrm_name` | Single line of text | Primary name; used as `labelFieldName` |
| `xrm_color` | Single line of text | Optional hex colour, e.g. `#2E7D32` |
| `xrm_icon` | Single line of text | Optional Fluent icon name or image URL |
| `xrm_category` | Single line of text | Optional category for grouping in picker |
| `xrm_description` | Multi-line text | Optional tooltip / subtitle |
| `xrm_sortorder` | Whole number | Optional display order |

### 2 – Create the M2M relationship

In the **account** table (or whichever host entity you want), add a **Many-to-Many relationship** to `xrm_keyword`. Note:

- **Relationship name** → `relationshipName` property
- **Navigation property** on the host entity → `navigationPropertyName` property
- **Entity set name** of `xrm_keyword` → `relatedEntitySetName` property

### 3 – Add the control to a form

1. Open the form editor for the host entity.
2. Add a **text field** as the bound column (the control itself does not write to this field — it manipulates the M2M relationship directly).
3. In the **Controls** tab, add **EntityKeywordControl**.
4. Fill in all required manifest properties.

---

## Keyword Record Metadata

### Colour

Set `colorFieldName` to a field that contains a CSS hex string such as `#E8A202`.  
The control automatically derives a tinted background and border at reduced opacity.

### Icon

Set `iconFieldName` to a field containing either:

- An **absolute image URL** (`https://...` or paths containing `/` or `.`)
- A **Fluent icon stub** (any other string) — the first character is displayed as a visual badge

### Grouping

Set `groupFieldName` to a choice or text field.  Keywords sharing the same group value are displayed under a labelled section header inside the picker dialog.

---

## Development Notes

### Runtime-only APIs

Two WebAPI methods used by this control (`associateRecord` and `disassociateRecord`) are available in the Dataverse runtime but are not yet declared in the published `@types/powerapps-component-framework` typings.  They are called through a typed `any` cast with a comment marking the workaround.  Remove the cast once the official typings are updated.

Similarly, `context.page.entityId` / `context.page.entityTypeName` are available at runtime in model-driven apps but absent from the published types.

### Fluent UI v9

The control uses `@fluentui/react-components` (v9) exclusively — no dependency on the older `@fluentui/react` (v8) package.  This keeps the bundle lean and forward-compatible with future Power Apps Fluent v9 adoption.

---

## Testing

Tests are written with [Jest](https://jestjs.io/) and [React Testing Library](https://testing-library.com/docs/react-testing-library/intro/).

```bash
npm test
```

The test suite covers:

- Loading / spinner state
- Empty state (no keywords, editing disabled)
- Tag rendering from associated records
- Add keyword button visibility based on `allowAssociate`
- Overflow badge at `maxVisibleTags` limit and expand behaviour
- `disassociateRecord` WebAPI call on tag dismiss
- Error message display on WebAPI failure
- Picker panel: search filtering, already-selected exclusion, multi-select and apply
- Standalone tag: label, dismiss, read-only mode, tooltip

---

## License

MIT
