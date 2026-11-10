# Excel Financial & Operational Tracker

A dynamic, multi-tab operational and financial tracking automation built in Microsoft Excel. Designed to manage corporate accounts, sales/purchase transactions, VAT calculations, and real-time ledger reconciliations efficiently.

## 🚀 Key Features & Structure

The workbook consists of three interconnected, dynamic sheets:

1. **`Cari_Listesi` (Client/Supplier Master Directory):**
   - Centralized database storing core partner details (Company Name, Type, Contact Info, City).
   - Utilizes Excel Tables (`Ctrl + T`) for automatic range expansion and dynamic structuring.

2. **`Hareketler` (Transaction Ledger):**
   - Records daily financial activities (Invoices, Waybills, Sales, Purchases, VAT rates).
   - **Automated Lookups:** Automatically retrieves company names from the master list using **`XLOOKUP` (`ÇAPRAZARA`)** and **`VLOOKUP` (`DÜŞEYARA`)**.
   - **Dynamic Calculations:** Computes total transaction amounts incorporating VAT multipliers `(1 + rate)`.

3. **`Cari_Mutabakat_Ozet` (Reconciliation Summary Dashboard):**
   - **Multi-Criteria Aggregation:** Automatically calculates total sales, purchases, and cash flows per client using **`SUMIF` (`ETOPLA`)** and **`SUMIFS` (`ÇOKETOPLA`)**.
   - **Real-Time Balances:** Dynamically computes remaining account balances and tracks financial/risk status.

## 🛠️ Excel Functions & Tools Used
- **Excel Tables (`Ctrl + T`)** for structured data management.
- **`XLOOKUP` / `VLOOKUP`** for relational data retrieval.
- **`SUMIF` & `SUMIFS`** for conditional financial aggregation.
- **`IF` Logic** for automated risk and status evaluation.

## 👤 Author
Developed as part of a professional data analysis and operational management portfolio.
