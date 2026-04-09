Local JSON files are loaded before Google Sheets when `CONFIG.DATA_SOURCE = 'local_first'`.

Supported format:

```json
[
  ["Column 1", "Column 2"],
  ["Value A", "Value B"]
]
```

You can also use:

```json
{
  "values": [
    ["Column 1", "Column 2"],
    ["Value A", "Value B"]
  ]
}
```

Files used by the app:

- `Data/select-fund.json`
- `Data/thai-annualized.json`
- `Data/thai-calendar.json`
- `Data/master-annualized.json`
- `Data/master-calendar.json`

If a file is empty, the page will show no rows.
