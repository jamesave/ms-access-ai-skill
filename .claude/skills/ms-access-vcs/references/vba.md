# VBA Reference

> Cross-reference: [Forms Critical Rules](forms.md#critical-rules) apply — especially Rule 2 (event wiring). See also [Conditional Formatting](conditional-formatting.md) for VBA `FormatConditions` patterns.

## VBA Modules (modules/*.bas)

- Start with `Attribute VB_Name = "ModuleName"`
- Use `#If VBA7 Then` / `#Else` / `#End If` for API declarations (required for 64-bit Access)
- `Option Compare Database` is standard

## VBA Code-Behind Files (.cls)

Form and report code-behind files (`.cls`) are created when `SplitLayoutFromVBA=true`. These files must **NOT** include the `VERSION 1.0 CLASS` header block or `Attribute VB_Name`. VCS handles class registration internally during import — including these causes import errors.

**Correct `.cls` format:**
```vba
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpen_Click()
    DoCmd.OpenForm "fDetail"
End Sub
```

**Wrong — do NOT include:**
```vba
VERSION 1.0 CLASS          <- WRONG: remove this entire block
BEGIN                       <- WRONG
  MultiUse = -1  'True     <- WRONG
END                         <- WRONG
Attribute VB_Name = "Form_fMyForm"  <- WRONG: remove this line
Attribute VB_GlobalNameSpace = False
...
```

The `VERSION 1.0 CLASS` block and `Attribute VB_Name` are standard VBA class file headers used by the VB editor, but VCS strips them on export and does not expect them on import.
