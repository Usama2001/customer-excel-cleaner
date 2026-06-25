# Customer Excel Cleaner

Static browser app for cleaning customer import workbooks.

## Current customer rules

- Export the Joblogic customer template column order with the instruction row.
- Leave `ID` blank and copy the source customer `ID` into `External ID`.
- Normalize the `Name` column by trimming and collapsing repeated whitespace.
- Clean `Address1`, `Address2`, `Address3`, `Address4`, and `Post Code` by splitting comma-separated address parts, removing duplicate address fragments, compacting blank address lines, moving UK postcodes into `Post Code`, and removing single-number-only address/postcode values such as `0`.
- Show distinct `Customer Type` values with counts.
- Clean `Tags` by trimming comma-separated values and removing duplicates.
- Blank `Invoice Payment Due` when the value is `0`.
- Check that the number of notes split by `<Note_Separator>` matches the comma-separated count in `Note Types`, without changing either column.

## Local run

Open `index.html` in a browser, or serve the folder with any static web server.

```powershell
python -m http.server 8080
```

## Deployment

The app is ready for GitHub Pages because it is static HTML, CSS, and JavaScript.
