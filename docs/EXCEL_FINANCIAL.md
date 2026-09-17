# Financial calculation compatibility

The calculator implements PMT, PV, FV, NPER, IPMT, PPMT, SLN, SYD, EFFECT and NOMINAL. These cover constant-period loan payments, present/future value, term length, interest/principal components, two depreciation methods and nominal/effective rate conversion. `financial.ts` keeps these algorithms independent of the spreadsheet UI.

## Verified behavior

`tests/fixtures/financial-input.json` contains 170 authored cases evaluated in Windows Excel 16.0 build 4627, locale 1031. `financial-oracle.json` records their values/errors, formulas and the exact fixture SHA-256. The fixture uses integer numeric strings where coercion is tested, avoiding assumptions about decimal separators in the installed German Excel.

The cases include zero and negative interest, fractional/negative periods where applicable, optional future values, beginning/end payment timing, invalid periods, zero divisors, direct/reference numeric text, logical values, blanks and propagated errors. Payment timing treats a nonzero argument as beginning-of-period in the recorded cases. EFFECT/NOMINAL truncate period counts and reject logical arguments; other tested financial functions accept logical numeric coercion.

Calculations use `log1p` and `expm1` to reduce cancellation in growth factors. The native comparison tolerance is `1e-10 * max(1, abs(expected))`; this is a numerical comparison, not bitwise equality or proof of every extreme floating-point case. Invalid counts/domains fail explicitly, and unsupported dependencies remain unsupported.

`src/financial.test.ts` compares all native cases, binds the receipt to the fixture and checks argument counts and unsupported dependency propagation. `tests/financial.spec.ts` imports deliberately stale loan caches, changes the interest rate to zero, checks payment/interest/principal and dependent totals, exercises undo/redo, reloads saved state and verifies the downloaded XLSX caches.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-formulas.ps1 -InputPath tests/fixtures/financial-input.json -OutputPath .local/financial-oracle.json
npm.cmd run test -- src/financial.test.ts
npm.cmd run test:e2e -- tests/financial.spec.ts
```

## Remaining scope

RATE, IRR/XIRR, NPV/XNPV, bond/security functions, cumulative financial functions, dynamic arrays, all locale coercions and extreme-rate precision remain incomplete. This implementation does not complete Excel's financial inventory or desktop parity.

Primary references: [Microsoft PV](https://support.microsoft.com/en-us/excel/functions/pv-function), [PMT](https://support.microsoft.com/en-us/excel/functions/pmt-function), [IPMT](https://support.microsoft.com/en-us/excel/functions/ipmt-function) and [EFFECT](https://support.microsoft.com/en-us/excel/functions/effect-function). The native receipt supplies independent results beyond the examples in those references.
