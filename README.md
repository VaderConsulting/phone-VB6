# phone

CSC VB6 HRMF→Exchange phonebook transfer (`HRMF2MSX1.exe` / `Hrmf2msx.vbp`): pushes HR/phonebook changes into Microsoft Exchange (optional AUTO mode), with companion `HRMFOptions` for registry options. Open `Hrmf2msx/Hrmf2msx.vbp` in the VB6 IDE.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

_Note: original OneDrive LastWriteTime values were wiped to 2026-08-27 by a zip transfer; date above uses best available evidence (headers/copyright where helpful)._

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `HRMFOptions` (`HRMFOptions/HRMFOptions.vbp`) | VB6 | WinForms exe | HRMFOptions |
| `HRMF2MSX` (`Hrmf2msx/Hrmf2msx.vbp`) | VB6 | WinForms exe | HRMF2MSX |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `HRMFOptions/HRMFOptions.vbp`
- `Hrmf2msx/Hrmf2msx.vbp`

## Requirements

- Visual Basic 6.0 IDE
- Registered OCX/DLL dependencies referenced by the `.vbp` (may need to be installed separately):
  - `TABCTL32.OCX`

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/Old/phone`.
Company names in project files: Computer Sciences Corporation.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
