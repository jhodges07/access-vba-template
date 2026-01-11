# access-vba-template
Starter template for Microsoft Access VBA projects using GitHub as the source of truth

## Access VBA ↔ GitHub Workflow

This repository treats GitHub as the **source of truth** for all VBA code.

### Rules

- `.accdb` / `.mdb` files are **never committed**
- All VBA objects are exported as text files
- Git tracks **code**, not binaries

### Folder Mapping

| Access Object | Folder |
|--------------|--------|
| Standard Modules (.bas) | `/src/modules` |
| Class Modules (.cls) | `/src/classes` |
| Forms (.frm + .frx) | `/src/forms` |
| Queries (SQL text) | `/src/queries` |
| Reports | `/src/reports` |
| JSON configs | `/src/json` |

### Development Flow

1. Develop inside Microsoft Access
2. Export VBA objects to `/src`
3. Commit changes to GitHub
4. Re-import from `/src` when needed

Access is a **runtime container**.  
GitHub is the **system of record**.

