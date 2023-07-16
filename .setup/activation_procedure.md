# Activate dev environment

1. Activate Node.js virtual environment

    For bash:

    ```bash
    .env\Scripts\activate
    ```

    For powershell:

    ```powershell
    .env\Scripts\Activate.ps1
    ```

2. Log in to Google account

    ```powershell
    clasp login
    ```

---

## Useful powershell commands

Check list of ignored files

```powershell
clasp status > .temp\out.txt
```

Push project changes to an App Script project

```powershell
clasp push
```

Delete metadata from `demo.xlsx` file

```powershell
demo\delete_metadata_from_office_file_in_folder.ps1
```