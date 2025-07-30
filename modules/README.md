# GST Tax Invoice System - Refactored VBA Modules

This directory contains the refactored VBA code for the GST Tax Invoice System. The original monolithic script has been broken down into six well-organized modules to improve readability, maintainability, and scalability.

## How to Use These Modules

To implement the new structure in your Excel workbook, follow these steps:

1.  **Open the VBA Editor** in Excel by pressing `Alt + F11`.
2.  **Remove the Old Code**: If you have the previous version of the code in a single module, it is recommended to remove it to avoid conflicts.
3.  **Import the New Modules**:
    *   In the VBA Editor, right-click on the "Project Explorer" pane (usually on the left).
    *   Select **Import File...**.
    *   Navigate to this `modules` directory.
    *   Import each of the `.bas` files one by one.

Once all six modules are imported, the system is ready to use.

## How to Run the System

The main entry points for the system are located in the **`Main_Setup.bas`** module. To get started, you can run one of the following macros:

*   **`QuickSetup()`**: This is the recommended macro for first-time users. It creates all necessary worksheets (`GST_Tax_Invoice_for_interstate`, `Master`, `warehouse`) without setting up data validation, making it a fast and simple way to get started.
*   **`StartGSTSystem()`**: This macro provides the complete setup, including the creation of all worksheets and the configuration of all data validation rules and dropdowns.

You can run these macros by pressing `Alt + F8` in Excel, selecting the desired macro from the list, and clicking "Run".

## Module Descriptions

Each module has a specific responsibility, ensuring a clear separation of concerns.

### 1. `Main_Setup.bas` (Main Module)

*   **Purpose**: This is the primary module for system initialization and user interaction. It contains the main subroutines that users will run to set up the invoice system.
*   **Key Functions**: `StartGSTSystem()`, `QuickSetup()`, `ShowAvailableFunctions()`.

### 2. `Module_InvoiceStructure.bas`

*   **Purpose**: Handles the creation, formatting, and layout of the main invoice worksheet. It is responsible for the visual structure of the invoice, including headers, footers, and cell formatting.
*   **Key Functions**: `CreateInvoiceSheet()`, `SetupTaxCalculationFormulas()`, `AutoPopulateInvoiceFields()`.

### 3. `Module_InvoiceEvents.bas`

*   **Purpose**: Contains all the event handlers for the buttons on the invoice sheet. This module manages all user interactions, such as saving an invoice, creating a new one, or printing to PDF.
*   **Key Functions**: `SaveInvoiceButton()`, `NewInvoiceButton()`, `PrintAsPDFButton()`, `AddNewItemRowButton()`.

### 4. `Module_Warehouse.bas`

*   **Purpose**: Manages all data-related operations for the `warehouse` sheet. This includes creating the sheet, managing customer and HSN code data, and setting up data validation and dropdown lists for the invoice.
*   **Key Functions**: `CreateWarehouseSheet()`, `SetupDataValidation()`, `GetCustomerDetails()`, `GetHSNDetails()`.

### 5. `Module_Master.bas`

*   **Purpose**: Responsible for all operations related to the `Master` sheet, which stores a record of all saved invoices. It also manages the automatic invoice numbering system.
*   **Key Functions**: `CreateMasterSheet()`, `GetNextInvoiceNumber()`.

### 6. `Module_Utilities.bas`

*   **Purpose**: A collection of shared helper functions that are used by other modules. This includes functions for checking if a worksheet exists, converting numbers to words, and cleaning text.
*   **Key Functions**: `WorksheetExists()`, `NumberToWords()`, `CleanText()`.