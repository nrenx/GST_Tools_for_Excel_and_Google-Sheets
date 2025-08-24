always use byterover-retrieve-knowledge tool to get the related context before any tasks 
always use byterover-store-knowledge to store all the critical informations after sucessful tasks

# GST Tax Invoice System – AI Coding Agent Instructions

## Essential Workflow

- **Always use `byterover-retrieve-knowledge`** before any task to gather relevant project context.
- **Always use `byterover-store-knowledge`** after completing a task to persist critical implementation details, patterns, and commands.

## Project Architecture

- **Modular VBA System**: Codebase is split into core modules (`1_Main_Setup.bas` to `6_Module_Utilities.bas`), button modules (`7_AddCustomerToWarehouseButton.bas` to `14_PDFUtilities.bas` in `button/`), and event/tax modules (`15_DataPopulation.bas` to `19_DataCleanup.bas` in `invoice_events/`).
- **Import Order Matters**: When importing into Excel VBA, always import core modules first, then button modules, then invoice_events modules. See `/README.md` and subfolder READMEs for exact order.
- **Main Entry Points**: Initialization is via `QuickSetup()` or `StartGSTSystem()` in `1_Main_Setup.bas`.

## Key Patterns & Conventions

- **Button Logic**: All button-related code is in `button/`. Each button module exposes a single main function (e.g., `AddCustomerToWarehouseButton`) for assignment to Excel UI buttons. Use `ButtonManagement.bas` to create/remove buttons programmatically.
- **Event & Tax Logic**: All worksheet event handlers and tax calculation logic are in `invoice_events/` modules. Functions are grouped by logical concern (data population, events, tax, display, cleanup).
- **Cross-Module Calls**: Button and event modules depend on core modules for data validation, tax calculations, and worksheet management. Maintain original function signatures for compatibility.
- **No Functional Changes in Refactoring**: All refactoring is organizational—functionality, error handling, and signatures are preserved for backward compatibility.

## Integration & Dependencies

- **macOS Compatibility**: PDF export and print functions are designed to work on macOS.
- **Dynamic Worksheet Creation**: System creates and formats all required sheets (`GST_Tax_Invoice`, `Master`, `warehouse`) on setup.
- **Dropdowns/Data Validation**: All dropdowns are populated from the `warehouse` sheet; validation setup is in `Module_Warehouse.bas`.

## Examples

```vba
' Assigning a button function
Call CreateButtonAtCell(ws, "R7", "Save Customer to Warehouse", "AddCustomerToWarehouseButton")

' Running setup
Call QuickSetup()
Call StartGSTSystem()

' Data population and event handling
Call AutoPopulateInvoiceFields(ws)
Call HandleSaleTypeChange(ws, changedRange)
Call SetupTaxCalculationFormulas(ws)
Call UpdateMultiItemTaxCalculations(ws)
```

## References

- See `/README.md`, `button/README.md`, and `invoice_events/README.md` for detailed module breakdowns, import order, and integration notes.

---
**Always use byterover tools for context retrieval and knowledge storage as your first and last step in any workflow.**[byterover-mcp]

# important 
always use byterover-retrieve-knowledge tool to get the related context before any tasks 
always use byterover-store-knowledge to store all the critical informations after sucessful tasks