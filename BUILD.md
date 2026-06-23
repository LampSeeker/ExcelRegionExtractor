# PyPI Release

## 1. Clean old build outputs

```powershell
Remove-Item -Recurse -Force build, dist, src/*.egg-info -ErrorAction SilentlyContinue
```

## 2. Run tests

```powershell
pytest
```

## 3. Bump version

Update `version` in `pyproject.toml`.

PyPI does not allow re-uploading the same version.

## 4. Build wheel and sdist

Install release tools once if needed:

```powershell
python -m pip install build twine
```

Build:

```powershell
python -m build
```

Outputs are written to `dist/`:

```text
dist/
  excel_region_extractor-*.whl
  excel_region_extractor-*.tar.gz
```

## 5. Check package metadata

```powershell
python -m twine check dist/*
```

## 6. Test the built wheel locally

```powershell
python -m pip install --force-reinstall dist/excel_region_extractor-*.whl
python -c "from excel_info_region import load_config; print(load_config()['extract_chart_images'])"
excel-regions --workbook examples/synthetic_demo.xlsx --out outputs/demo --no-overlay
```

## 7. Upload to PyPI

Run:

```powershell
python -m twine upload dist/*
```

When `twine` prompts for credentials, type:

```text
username: __token__
password: <your PyPI API token, starts with pypi->
```

Do not paste the block above into PowerShell. It is the interactive input for the prompt.

Non-interactive option:

```powershell
$env:TWINE_USERNAME = "__token__"
$env:TWINE_PASSWORD = "<your PyPI API token>"
python -m twine upload dist/*
Remove-Item Env:\TWINE_USERNAME
Remove-Item Env:\TWINE_PASSWORD
```

## 8. Verify install from PyPI

Use a clean environment if possible:

```powershell
python -m pip install --upgrade excel-region-extractor
python -c "from excel_info_region import load_config; print(load_config()['extract_chart_images'])"
excel-regions --help
```

## 403 Forbidden

`403 Forbidden` means PyPI rejected the upload credentials or project ownership.

Check these first:

```powershell
python -m twine upload --verbose dist/*
```

Common causes:

- The token is wrong, expired, or copied with extra spaces.
- A project-scoped token was used before the project exists. Use an account-scoped PyPI API token for the first upload.
- The package name in `pyproject.toml` is already owned by another PyPI account.
- The same version was already uploaded. Bump `version` in `pyproject.toml` and rebuild.

Package name and version come from:

```toml
[project]
name = "excel-region-extractor"
version = "0.1.2"
```

## Git Ignore

Release outputs and local secrets are ignored by `.gitignore`:

```text
*.egg-info/
build/
dist/
pip-wheel-metadata/
outputs/
.pypirc
pypi-token.txt
```
