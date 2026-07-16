# Virelo Release Checklist

## 1. Prepare a Clean Versioned Source

1. Update `APP_VERSION` in `virelo/app/config.py`.
2. Synchronize the npm package and lock records:

   ```powershell
   $Version = Read-Host "Release version"
   cd frontend
   npm version $Version --no-git-tag-version
   cd ..
   ```

3. Review and commit only the intended release files:

   ```powershell
   git diff --check
   git add virelo/app/config.py frontend/package.json frontend/package-lock.json
   git commit -m "chore(release): prepare v$Version"
   ```

4. Confirm that the complete Git working tree is clean, including no untracked files:

   ```powershell
   git status --short
   ```

5. Recreate and audit the locked environments:

   ```powershell
   scripts/bootstrap.ps1 -Recreate
   .venv\Scripts\python.exe -m pip check
   cd frontend
   npm audit --audit-level=high
   npm run lint
   npm run format:check
   npm test
   cd ..
   ```

6. Review the exact bundled dependency set and include every third-party notice and license text
   required for the intended distribution channel. The repository's MIT license does not replace
   the obligations of Qt, PySide6, or other redistributed libraries.

## 2. Build and Sign in the Correct Order

Build from an unprivileged terminal. Retain the complete command output as the build log.

```powershell
scripts/clean.ps1
scripts/build-app.ps1
```

For authenticated direct distribution, sign the final application executable with an approved
code-signing identity and RFC 3161 timestamp. Keep certificate selection and credentials outside
the repository and logs.

```powershell
signtool sign /fd SHA256 /tr $TimestampUrl /td SHA256 /a dist\Virelo\Virelo.exe
signtool verify /pa /all /v dist\Virelo\Virelo.exe
scripts/write-bundle-checksums.ps1
scripts/build-installer.ps1 -SkipAppBuild
signtool sign /fd SHA256 /tr $TimestampUrl /td SHA256 /a installer\dist\VireloSetup.exe
signtool verify /pa /all /v installer\dist\VireloSetup.exe
scripts/write-release-checksums.ps1
scripts/verify-release.ps1 -RequireSignature
```

For an intentionally unsigned internal candidate, use this separate sequence:

```powershell
scripts/clean.ps1
scripts/build-installer.ps1
scripts/verify-release.ps1
```

Retain the reported `NotSigned` warnings in the acceptance evidence. Do not present an unsigned
build as a signed public release.

## 3. Inspect Static Release Evidence

Confirm that verification reports the expected commit and version, a clean source build, exact
input and frontend hashes, a complete bundle checksum inventory, valid artifact checksums, PE
version metadata, packaged smoke success, and the intended signature status.

Record at least:

- The source commit and tag.
- Python, pip, PyInstaller, Node.js, npm, and PowerShell versions.
- `dist/Virelo/.release.json`.
- `dist/Virelo/bundle-files.sha256`.
- `installer/dist/CHECKSUMS.sha256`.
- Authenticode verification output.
- The CI run URL and result.

## 4. Exercise the Installed Lifecycle

Use a clean supported Windows machine or disposable virtual machine without the development
runtime on `PATH`. Static verification does not replace these tests.

| Initial version and state | Action or injected failure | Expected installed version | Expected retained state | Expected recovery result |
|---|---|---|---|---|
| No installation | Install, launch, and exercise the workflows below. | Candidate | New default settings. | The installed application remains usable after restart. |
| Immediately previous release with customized settings | Upgrade to the candidate. | Candidate | Settings and folder-view recovery backups remain available. | The candidate launches and repeats the workflows. |
| Previous release with `_internal/removed-release-probe.txt` | Upgrade to the candidate. | Candidate | User settings remain available. | The stale probe is absent before launch. |
| Candidate running in the tray | Start the installer. | Candidate after the app is closed. | Settings and backups remain available. | `AppMutex` prevents an unsafe in-place replacement. |
| Previous release with an intentionally aborted candidate installation | Follow the documented manual recovery by reinstalling the previous installer. | Previous release | Restore or retain the copied prior state. | The previous release launches and repeats the workflows. |
| Fresh candidate installation | Uninstall, verify installer-owned removal, reinstall, launch, and repeat the workflows. | Candidate after reinstall | Settings and recovery backups follow the retention contract. | Reinstall is usable. |
| Upgraded candidate installation | Uninstall, verify installer-owned removal, reinstall, launch, and repeat the workflows. | Candidate after reinstall | Settings and recovery backups follow the retention contract. | Reinstall is usable. |

The machine-wide uninstaller removes only installer-owned files. It retains every account's
settings, logs, folder-view recovery backups, and startup shortcut. Turn off **Run at Startup**
from each owning account before uninstalling, or remove that account's `Virelo.lnk` manually from
`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup` afterward. Alternate-credential UAC
scenarios are outside the supported acceptance boundary because the elevated process runs under
the credential account.

Inno Setup does not provide an automatic application rollback here. Preserve the prior installer
and document manual recovery instead of claiming rollback support.

### Required Workflows

1. Snap a normal foreground window and verify centering against the visible work area, then
   restore it. Repeat on monitors with different scale factors and negative coordinates. Center
   Google Drive and LightBulb, verify that neither window changes size, and restore each window
   to its exact original position.
2. Enable Explorer column auto-size, navigate between folders in Details view, and verify that
   columns fit and that the idle worker backs off to a one-second cadence with negligible CPU use.
3. Apply Details as the default folder view and confirm the displayed backup path. After active
   file operations finish, manually restart File Explorer or sign out, then verify representative
   folder types. Reset folder views and confirm recovery behavior.
4. Save settings, restart Virelo, and verify the saved values, tray behavior, and startup shortcut.

Capture actual results, timestamps, screenshots or logs, and hashes for every lifecycle row.

## 5. Tag and Publish Deliberately

Tag the exact commit that produced the verified artifacts. If any source changed after the build,
discard the artifacts and restart the build and verification process from a clean commit.

```powershell
$Version = Read-Host "Release version"
git tag -s "v$Version" -m "Virelo v$Version"
git push origin main
git push origin "refs/tags/v$Version"
```

Use an annotated tag when signing is unavailable, and record that boundary. Publishing a GitHub
release and uploading artifacts are separate authorized actions. Attach the final signed installer,
checksum file, release notes, and acceptance evidence only after all required gates pass.
