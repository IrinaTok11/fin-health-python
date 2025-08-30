# Architecture (high-level)

```
┌────────────────┐      ┌────────────────────┐
│   Excel file   │      │  ratio_norms sheet │
│ (CWD, single)  │      │  norms/units/bias  │
└──────┬─────────┘      └──────────┬─────────┘
       │                            │
       ▼                            ▼
┌────────────────────────────────────────────┐
│            run/summary.py                  │
│  - Parse years/params/income/balance       │
│  - Compute 12 KPIs                         │
│  - Build Summary (norms, trend, change)    │
│  - Write Excel sheet + Word 3.2 (4 KPIs)   │
└───────────────────────┬────────────────────┘
                        │
                        ▼
           ┌─────────────────────────┐
           │  Excel Summary (12)     │
           ├─────────────────────────┤
           │ Word 3.2 (4 liquidity)  │
           └─────────────────────────┘
```

Principles: **reproducible**, **portable**, **opinionated defaults**.
