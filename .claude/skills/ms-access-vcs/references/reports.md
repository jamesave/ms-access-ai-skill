# Reports Reference

> See also [Conditional Formatting](conditional-formatting.md) for format conditions on report controls.

## Critical Rules

These rules apply when editing any report `.bas` file. Violations cause silent corruption or `Error 2128` on import.

1. **TabIndex values MUST be contiguous within each section (0, 1, 2, 3...).** Gaps cause `Error 2128` on import. When removing a control, renumber all subsequent TabIndex values in that section downward to fill the gap.

2. **VBA event procedures MUST be wired up in the `.bas` file.** Adding a `Private Sub Report_Open()` (or any event) to the `.cls` code-behind is **not sufficient** — Access will silently ignore it. You must also add the corresponding event property in the `.bas` layout file. The property goes on the **object that owns the event**:
   - **Report-level events** (`Report_Open`, `Report_Load`, etc.) → property on the report object: `OnOpen ="[Event Procedure]"`, `OnLoad ="[Event Procedure]"`, etc.
   - **Control events** (`txtName_AfterUpdate`, etc.) → property on the control block: `AfterUpdate ="[Event Procedure]"`, etc.
   - **Section events** (`Detail_Format`, etc.) → property on the section block: `OnFormat ="[Event Procedure]"`, etc.

   **Without this property, the VBA code compiles but never executes.**

3. **Never edit binary data blocks.** Leave `ImageData = Begin...End`, `RecSrcDt = Begin...End`, and other hex-encoded blocks untouched. Modifying these corrupts the object. For `ConditionalFormat` and `ConditionalFormat14` blocks specifically, see [Conditional Formatting](conditional-formatting.md) — these must be removed entirely and replaced with VBA code.

4. **Preserve Begin/End nesting exactly.** Every `Begin` must have a matching `End`. Mismatched nesting causes import failure.

5. **Control order matters.** Controls are rendered in file order (affects z-order). Don't reorder controls unless intentionally changing layering.

6. **Property values use specific formats:**
   - Numeric: `Width =1701` (units are twips: 1 inch = 1440 twips)
   - String: `Name ="Detail"` (always double-quoted)
   - Boolean flags: presence of property name implies true (e.g., `NotDefault`)
   - **String escaping uses backslash (`\"`), NOT VBA-style doubled quotes (`""`).**

7. **When removing a control block,** remove everything from `Begin ControlType` through its closing `End`, including any trailing `LayoutCached*` properties. Then renumber TabIndex values to fill the gap (Rule 1).

## Report Structure

Reports use the same binary text format as forms (`Begin Report` instead of `Begin Form`). The same container structure applies: exactly ONE `Begin...End` container block holds default control styles, break level definitions, and all sections.

### Report Section Keywords

Report sections use these keywords — note that Report Header/Footer reuse the `FormHeader`/`FormFooter` keywords:

| Section | Keyword | Name Property |
|---|---|---|
| Report Header | `FormHeader` | `Name ="ReportHeader"` |
| Page Header | `PageHeader` | `Name ="PageHeaderSection"` or `Name ="PageHeader"` |
| Group Header | `BreakHeader` | `Name ="GroupHeader0"` (0-indexed) |
| Detail | `Section` | `Name ="Detail"` |
| Group Footer | `BreakFooter` | `Name ="GroupFooter0"` or `Name ="GroupFooter1"` |
| Page Footer | `PageFooter` | `Name ="PageFooterSection"` or `Name ="PageFooter"` |
| Report Footer | `FormFooter` | `Name ="ReportFooter"` |

Note: Report Header/Footer sections use the **`FormHeader`/`FormFooter` keywords** (not `ReportHeader`/`ReportFooter`). Group Header/Footer sections use **`BreakHeader`/`BreakFooter`** (not `GroupHeader0`/`GroupFooter0`).

## Sorting and Grouping

Reports define grouping via `Begin BreakLevel` blocks inside the container, placed **after** default control styles and **before** sections. Each `BreakLevel` defines one sorting/grouping level.

```
    Begin                               <- Container
        Begin Label                     <- Default styles first
            ...
        End
        Begin TextBox
            ...
        End
        Begin BreakLevel                <- Group level definitions
            GroupHeader = NotDefault     <- Show group header section
            GroupFooter = NotDefault     <- Show group footer section
            ControlSource ="OrderID"    <- Field to group by
        End
        Begin FormHeader                <- Sections follow
            Name ="ReportHeader"
            ...
        End
        Begin PageHeader
            ...
        End
        Begin BreakHeader               <- Group header section
            Name ="GroupHeader0"
            Begin
                ...controls...
            End
        End
        Begin Section
            Name ="Detail"
            ...
        End
        Begin BreakFooter               <- Group footer section
            Name ="GroupFooter0"
            Begin
                ...controls...
            End
        End
        Begin PageFooter
            ...
        End
        Begin FormFooter
            Name ="ReportFooter"
        End
    End                                 <- Closes container
```

### BreakLevel Properties

- `ControlSource ="FieldName"` — the field to sort/group by (required)
- `GroupHeader = NotDefault` — show a group header section (optional; omit for sort-only levels)
- `GroupFooter = NotDefault` — show a group footer section (optional; omit for sort-only levels)

### Multiple Break Levels

Define multiple `Begin BreakLevel...End` blocks for multi-level sorting/grouping. Only levels with `GroupHeader`/`GroupFooter = NotDefault` produce visible sections. When multiple break levels exist, add `BreakLevel =N` (0-indexed) on the `BreakHeader`/`BreakFooter` sections to indicate which level they belong to. When there is only one break level, the `BreakLevel` property can be omitted.

Example with 4 break levels (only level 2 has visible header/footer):
```
        Begin BreakLevel
            ControlSource ="JobDate"
        End
        Begin BreakLevel
            ControlSource ="Driver"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Driver"
        End
        Begin BreakLevel
            ControlSource ="StoreArrivalTime"
        End
        ...
        Begin BreakHeader
            BreakLevel =2
            Name ="GroupHeader0"
            ...
        End
        ...
        Begin BreakFooter
            BreakLevel =2
            Name ="GroupFooter1"
            ...
        End
```

## Common Structural Errors

- Using `Begin Sorting` → `Expected: 'Begin'. Found: Sorting.` (`Sorting` is not valid syntax; use `BreakLevel` instead)
- Using `Begin GroupHeader0` or `Begin GroupFooter0` → `Expected: Object Type Name. Found: GroupHeader0.` (use `BreakHeader`/`BreakFooter` instead)
- Using `Begin ReportHeader` or `Begin ReportFooter` → use `Begin FormHeader`/`Begin FormFooter` with the appropriate `Name` property instead
- Placing `Begin Label` or sections directly inside `Begin Report` without the container `Begin` → `Expected: 'Begin'. Found: Label.`
