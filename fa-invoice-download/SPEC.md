# FA Invoice Download - Specification

## Purpose

Download XML invoices from the tax portal for a requested period and update a local control-ready review workbook.

## Skill folder layout

The skill now keeps helper files at the skill root, next to `SKILL.md`, instead of using a nested `scripts/` folder.

```txt
fa-invoice-download/
├─ SKILL.md
├─ SPEC.md
├─ icon.svg
└─ update_invoice_list.py
```

## Expected folder structure

```txt
<folderPath>/
├─ credential.txt or Credentials.txt
├─ Files/
└─ InvoiceList.xlsx
```

## Portal flow

- Read credentials from text file
- Login with browser automation
- Navigate to invoice XML export / download area
- Filter by requested date range
- Download all XML files
- Rename each file as `{invoiceSeries}_{invoiceNumber}.xml`
- When the portal UI lookup flow is broken, allow authenticated API fallback after successful login

## Workbook target

Workbook: `InvoiceList.xlsx`

Sheet: `Invoice_Tax_Lines`

### Review columns
- InvoiceDate
- InvoiceNo
- SellerCompany
- SellerTaxCode
- BuyerName
- BuyerTaxCode
- Description
- TaxRate
- AmountBeforeTax
- TaxAmount
- AmountAfterTax
- DetailsText
- Status
- Notes

### Control-ready columns
- InvoiceFormNo
- InvoiceSeries
- InvoiceNumber
- TaxAuthorityCode
- SigningDate
- InvoiceType
- InvoiceNature
- PaymentMethod
- SellerAddress
- BuyerAddress
- Currency
- ExchangeRate
- AmountInWords
- SourceFile

## Parsing strategy

The bundled root-level Python script `update_invoice_list.py` uses a heuristic XML parser:

1. Standard XML tag extraction when invoice fields are present in structured nodes.
2. When an invoice has multiple VAT rates, create one Excel row per VAT rate.
3. Repeat the shared invoice metadata identically across those rows.
4. Prefer VAT-summary nodes in XML; if unavailable, derive groups from invoice line/item nodes.
5. Fallback to total-level extraction when line-level data is missing.
6. Header upgrade when the workbook still uses an older column set.
7. Skip duplicate rows when the invoice already exists in Excel.

## Duplicate control

Default duplicate key:
- InvoiceNo + Description + TaxRate + AmountBeforeTax

## Output artifacts

- Updated `InvoiceList.xlsx`
- `invoice_update_summary.json`
- `invoice_parse_errors.csv` when parse issues exist

## Known limitations

- Portal UI may change.
- Captcha / OTP requires human intervention.
- XML schemas vary across invoice providers.
- Some invoices may only support summary extraction instead of detailed tax-line extraction.

## Control points

- Never log plaintext passwords.
- Never silently ignore parse failures.
- Keep reruns idempotent through strict dedupe behavior: existing invoices are skipped, not re-imported.
- Produce review-ready rows only; no accounting posting.
