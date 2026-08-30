# CHAEN POS v1.23

Product master is now read from the Google Sheet instead of being hard-coded.

## What changed

- `Recipe Master` (Google Sheets) is the single source of truth for products.
  The POS fetches the list at start-up from Apps Script `doGet(action=products)`
  and builds the order buttons from it.
- Adding a drink to the sheet — or changing a price — now shows up in the POS
  after a reload. No code change, no redeploy.
- Sizes and prices come from the sheet per row. The old "22oz = 16oz + P30"
  assumption is gone, so a product can have any set of sizes.
- Products are grouped into category tabs (Matcha / Milky / Frappe / Churros)
  driven by the `Category` column.
- The fetched list is cached in `localStorage`. If the sheet cannot be reached
  the POS uses the saved copy; if there is no saved copy either (first install
  while offline) it falls back to a built-in list so the register still works.
- The order screen shows when the product list was last synced, and has a
  manual "Refresh product list" button.

## P30 OFF rule

Now decided by category instead of a name blacklist:

    matcha drinks, 22oz only, Signature Matcha Latte excluded

Non-matcha products (Milky / Frappe / Churros) are never discountable.
Behaviour for the existing matcha drinks is unchanged. See `canDiscount()`
in `app.js`.

## Unchanged

- `doPost` (writing sales and expenses to the sheet) is untouched.
- History edit/delete, Close Day, Expense, backdated orders, multi-device
  same-day sync, English UI.
