"""Explorer default folder view management.

Makes Details view the default for every File Explorer folder type using
the same registry mechanism as LesFerch/WinSetView, reduced to a single
opinionated action:

1. Back up the affected HKCU registry keys to a structured JSON snapshot.
2. Delete the per-folder view caches (Bags and BagMRU in both hives) and
   the saved view defaults (Streams\\Defaults) so stale states cannot
   shadow the new defaults.
3. Copy HKLM FolderTypes to HKCU (Explorer prefers the HKCU copy) and
   force LogicalViewMode=Details on every TopViews entry.
4. Write Details-view bag entries for This PC, which has no FolderTypes
   GUID of its own.
5. Tell the user to restart Explorer or sign out so the shell drops its
   cached view state.

The module separates pure plan construction (testable everywhere) from
the Windows-only executor (winreg is imported lazily so unit tests can
import this module on any platform).
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import json
import logging
import os
from dataclasses import dataclass
from datetime import UTC, datetime

LOG = logging.getLogger("Virelo")

# Registry paths, all relative to HKCU unless noted otherwise.
SHELL_CLASSES = r"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell"
BAGS_KEY = SHELL_CLASSES + r"\Bags"
BAGMRU_KEY = SHELL_CLASSES + r"\BagMRU"
SHELL_LEGACY = r"Software\Microsoft\Windows\Shell"
BAGS_LEGACY_KEY = SHELL_LEGACY + r"\Bags"
BAGMRU_LEGACY_KEY = SHELL_LEGACY + r"\BagMRU"
STREAMS_DEFAULTS_KEY = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Streams\Defaults"
FOLDER_TYPES_KEY = r"Software\Microsoft\Windows\CurrentVersion\Explorer\FolderTypes"

# This PC has no FolderTypes entry; its view lives in a numbered bag.
THIS_PC_GUID = "{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}"
THIS_PC_PIDL = bytes.fromhex("14001F50E04FD020EA3A6910A2D808002B30309D0000")

# Details view constants (WinSetView SetViewValues index 1).
LOGICAL_VIEW_MODE_DETAILS = 1
MODE_DETAILS = 4
ICON_SIZE_DETAILS = 16
BAG_FFLAGS = 0x41200001
GROUP_BY_FMTID = "{B725F130-47EF-101A-A5F1-02608C9EEBAC}"
GROUP_BY_PID = 4

# Keys removed when clearing cached view state (HKCU-relative).
VIEW_CACHE_KEYS = (
    BAGMRU_KEY,
    BAGS_KEY,
    BAGMRU_LEGACY_KEY,
    BAGS_LEGACY_KEY,
    STREAMS_DEFAULTS_KEY,
    FOLDER_TYPES_KEY,
)

# Keys exported to the backup before anything is modified.
BACKUP_KEYS = VIEW_CACHE_KEYS
BACKUP_MANIFEST = "manifest.json"
BACKUP_DIGEST = "manifest.sha256"
BACKUP_PENDING_MARKER = ".pending"
BACKUP_SCHEMA_VERSION = 2
MAX_BACKUP_BYTES = 128 * 1024 * 1024
MAX_BACKUP_DEPTH = 64

# Registry kinds that can be represented safely by the structured snapshot.
# REG_LINK is intentionally excluded. These backups contain only ordinary
# Explorer view data, and restore is constrained to the fixed HKCU keys above.
_REGISTRY_DATA_KINDS = {
    "REG_NONE": "bytes",
    "REG_SZ": "str",
    "REG_EXPAND_SZ": "str",
    "REG_BINARY": "bytes",
    "REG_DWORD": "int",
    "REG_DWORD_BIG_ENDIAN": "int",
    "REG_MULTI_SZ": "str-list",
    "REG_RESOURCE_LIST": "bytes",
    "REG_FULL_RESOURCE_DESCRIPTOR": "bytes",
    "REG_RESOURCE_REQUIREMENTS_LIST": "bytes",
    "REG_QWORD": "int",
}


@dataclass(frozen=True)
class RegValue:
    """One registry value write, relative to HKCU."""

    key: str
    name: str
    kind: str  # "dword", "sz", or "binary"
    data: int | str | bytes


def this_pc_bag_values() -> list[RegValue]:
    """Build the value writes that force Details view for This PC."""
    values = [
        RegValue(BAGMRU_KEY, "NodeSlots", "binary", b"\x02"),
        RegValue(BAGMRU_KEY, "MRUListEx", "binary", bytes.fromhex("00000000ffffffff")),
        RegValue(BAGMRU_KEY, "0", "binary", THIS_PC_PIDL),
        RegValue(BAGMRU_KEY + r"\0", "NodeSlot", "dword", 1),
    ]
    for bag in (
        rf"{BAGS_KEY}\1\Shell\{THIS_PC_GUID}",
        rf"{BAGS_KEY}\1\ComDlg\{THIS_PC_GUID}",
    ):
        values.extend(
            [
                RegValue(bag, "FFlags", "dword", BAG_FFLAGS),
                RegValue(bag, "LogicalViewMode", "dword", LOGICAL_VIEW_MODE_DETAILS),
                RegValue(bag, "Mode", "dword", MODE_DETAILS),
                RegValue(bag, "GroupView", "dword", 1),
                RegValue(bag, "IconSize", "dword", ICON_SIZE_DETAILS),
                RegValue(bag, "GroupByKey:FMTID", "sz", GROUP_BY_FMTID),
                RegValue(bag, "GroupByKey:PID", "dword", GROUP_BY_PID),
            ]
        )
    return values


def top_view_values(top_view_key: str) -> list[RegValue]:
    """Build the value writes that force Details on one TopViews entry."""
    return [
        RegValue(top_view_key, "LogicalViewMode", "dword", LOGICAL_VIEW_MODE_DETAILS),
        RegValue(top_view_key, "IconSize", "dword", ICON_SIZE_DETAILS),
    ]


def backup_dir_name(now: datetime) -> str:
    """Return the timestamped directory name for a registry backup."""
    return now.strftime("view-backup-%Y%m%d-%H%M%S-%f")


# ----------------------------------------------------------------------------
# Windows-only executor (winreg imported lazily; safe to import cross-platform)
# ----------------------------------------------------------------------------


def _open_winreg():
    import winreg

    return winreg


def _delete_key_tree(winreg, root, path: str) -> None:
    """Recursively delete a registry key. Missing keys are not an error."""
    access = winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY
    try:
        key = winreg.OpenKey(root, path, 0, access)
    except FileNotFoundError:
        return
    try:
        while True:
            try:
                child = winreg.EnumKey(key, 0)
            except OSError as exc:
                if getattr(exc, "winerror", 259) == 259:
                    break
                raise
            _delete_key_tree(winreg, root, path + "\\" + child)
    finally:
        key.Close()
    try:
        winreg.DeleteKeyEx(root, path, winreg.KEY_WOW64_64KEY, 0)
    except FileNotFoundError:
        return
    except OSError:
        LOG.exception("Could not delete registry key HKCU\\%s", path)
        raise


def _copy_key_tree(winreg, src_root, src_path: str, dst_root, dst_path: str) -> None:
    """Recursively copy a registry key tree, preserving value types."""
    read = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
    write = winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
    with winreg.OpenKey(src_root, src_path, 0, read) as src:
        with winreg.CreateKeyEx(dst_root, dst_path, 0, write) as dst:
            index = 0
            while True:
                try:
                    name, data, kind = winreg.EnumValue(src, index)
                except OSError as exc:
                    if getattr(exc, "winerror", 259) == 259:
                        break
                    raise
                winreg.SetValueEx(dst, name, 0, kind, data)
                index += 1
        index = 0
        while True:
            try:
                child = winreg.EnumKey(src, index)
            except OSError as exc:
                if getattr(exc, "winerror", 259) == 259:
                    break
                raise
            _copy_key_tree(
                winreg, src_root, src_path + "\\" + child, dst_root, dst_path + "\\" + child
            )
            index += 1


def _write_values(winreg, values: list[RegValue]) -> None:
    """Apply and verify a list of RegValue writes under HKCU."""
    kinds = {
        "dword": winreg.REG_DWORD,
        "sz": winreg.REG_SZ,
        "binary": winreg.REG_BINARY,
    }
    access = winreg.KEY_READ | winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
    for value in values:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, value.key, 0, access) as key:
            expected_kind = kinds[value.kind]
            winreg.SetValueEx(key, value.name, 0, expected_kind, value.data)
            actual_data, actual_kind = winreg.QueryValueEx(key, value.name)
            if actual_kind != expected_kind or actual_data != value.data:
                raise RuntimeError(
                    f"Registry verification failed for HKCU\\{value.key}\\{value.name}."
                )


def _force_details_on_folder_types(winreg) -> int:
    """Force Details on every TopViews entry of the HKCU FolderTypes copy.

    Returns the number of TopViews entries updated.
    """
    read = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
    updated = 0
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, FOLDER_TYPES_KEY, 0, read) as folder_types:
        index = 0
        while True:
            try:
                type_guid = winreg.EnumKey(folder_types, index)
            except OSError as exc:
                if getattr(exc, "winerror", 259) == 259:
                    break
                raise
            index += 1
            top_views = rf"{FOLDER_TYPES_KEY}\{type_guid}\TopViews"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, top_views, 0, read) as views:
                    view_index = 0
                    view_guids = []
                    while True:
                        try:
                            view_guids.append(winreg.EnumKey(views, view_index))
                        except OSError as exc:
                            if getattr(exc, "winerror", 259) == 259:
                                break
                            raise
                        view_index += 1
            except FileNotFoundError:
                continue
            for view_guid in view_guids:
                _write_values(winreg, top_view_values(rf"{top_views}\{view_guid}"))
                updated += 1
    return updated


def _key_exists(winreg, key: str) -> bool:
    try:
        winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY
        ).Close()
        return True
    except FileNotFoundError:
        return False


def _backup_root() -> str:
    """Return the directory that contains Explorer view backups."""
    return os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "Virelo")


def _create_backup_directory(now: datetime | None = None) -> str:
    """Create and return a collision-resistant backup directory."""
    base = _backup_root()
    os.makedirs(base, exist_ok=True)
    timestamp = now or datetime.now()
    stem = backup_dir_name(timestamp)
    for suffix in range(1000):
        name = stem if suffix == 0 else f"{stem}-{suffix:03d}"
        target = os.path.join(base, name)
        try:
            os.mkdir(target)
            with open(os.path.join(target, BACKUP_PENDING_MARKER), "x", encoding="ascii"):
                pass
            return target
        except FileExistsError:
            continue
    raise RuntimeError("Could not create a unique Explorer view backup directory.")


def _registry_kind_maps(winreg) -> tuple[dict[int, str], dict[str, int]]:
    """Return canonical mappings for registry kinds supported by snapshots."""
    by_number: dict[int, str] = {}
    by_name: dict[str, int] = {}
    for name in _REGISTRY_DATA_KINDS:
        number = getattr(winreg, name, None)
        if isinstance(number, int):
            by_number.setdefault(number, name)
            by_name[name] = number
    return by_number, by_name


def _encode_registry_data(kind_name: str, data) -> dict:
    """Encode one registry value without losing its native Python type."""
    expected_type = _REGISTRY_DATA_KINDS.get(kind_name)
    value: str | int | list[str]
    if expected_type == "bytes" and isinstance(data, bytes):
        value = base64.b64encode(data).decode("ascii")
    elif expected_type == "str" and isinstance(data, str):
        value = data
    elif expected_type == "int" and isinstance(data, int) and not isinstance(data, bool):
        value = data
    elif (
        expected_type == "str-list"
        and isinstance(data, list)
        and all(isinstance(item, str) for item in data)
    ):
        value = data
    else:
        raise RuntimeError(f"Unsupported data for registry kind {kind_name}.")
    return {"type": expected_type, "value": value}


def _decode_registry_data(kind_name: str, encoded) -> bytes | str | int | list[str]:
    """Validate and decode one registry value from a snapshot."""
    expected_type = _REGISTRY_DATA_KINDS.get(kind_name)
    if expected_type is None or not isinstance(encoded, dict):
        raise RuntimeError("Backup contains an unsupported registry value kind.")
    if set(encoded) != {"type", "value"} or encoded.get("type") != expected_type:
        raise RuntimeError("Backup contains an invalid registry value encoding.")

    value = encoded.get("value")
    if expected_type == "bytes":
        if not isinstance(value, str):
            raise RuntimeError("Backup contains invalid binary registry data.")
        try:
            return base64.b64decode(value, validate=True)
        except (binascii.Error, ValueError, TypeError) as exc:
            raise RuntimeError("Backup contains invalid binary registry data.") from exc
    if expected_type == "str":
        if isinstance(value, str):
            return value
    elif expected_type == "int":
        if isinstance(value, int) and not isinstance(value, bool):
            return value
    elif expected_type == "str-list":
        if isinstance(value, list) and all(isinstance(item, str) for item in value):
            return value
    raise RuntimeError("Backup contains invalid registry value data.")


def _snapshot_key_tree(winreg, root, path: str, depth: int = 0) -> dict:
    """Capture one registry tree using only the worker process's HKCU access."""
    if depth > MAX_BACKUP_DEPTH:
        raise RuntimeError("Explorer registry tree is too deeply nested to back up safely.")
    read = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
    by_number, _ = _registry_kind_maps(winreg)
    values = []
    children = []
    with winreg.OpenKey(root, path, 0, read) as key:
        index = 0
        while True:
            try:
                name, data, kind = winreg.EnumValue(key, index)
            except OSError as exc:
                if getattr(exc, "winerror", 259) == 259:
                    break
                raise
            kind_name = by_number.get(kind)
            if kind_name is None:
                raise RuntimeError(f"HKCU\\{path} contains unsupported registry kind {kind}.")
            values.append(
                {
                    "name": name,
                    "kind": kind_name,
                    "data": _encode_registry_data(kind_name, data),
                }
            )
            index += 1

        index = 0
        while True:
            try:
                children.append(winreg.EnumKey(key, index))
            except OSError as exc:
                if getattr(exc, "winerror", 259) == 259:
                    break
                raise
            index += 1

    subkeys = [
        {
            "name": child,
            "tree": _snapshot_key_tree(winreg, root, path + "\\" + child, depth + 1),
        }
        for child in children
    ]
    return {"values": values, "subkeys": subkeys}


def _validate_snapshot_tree(tree, depth: int = 0) -> None:
    """Fail closed on malformed or ambiguous snapshot trees."""
    if depth > MAX_BACKUP_DEPTH:
        raise RuntimeError("Backup registry tree is too deeply nested.")
    if not isinstance(tree, dict) or set(tree) != {"values", "subkeys"}:
        raise RuntimeError("Backup contains an invalid registry tree.")
    values = tree.get("values")
    subkeys = tree.get("subkeys")
    if not isinstance(values, list) or not isinstance(subkeys, list):
        raise RuntimeError("Backup contains an invalid registry tree.")

    value_names: set[str] = set()
    for value in values:
        if not isinstance(value, dict) or set(value) != {"name", "kind", "data"}:
            raise RuntimeError("Backup contains an invalid registry value.")
        name = value.get("name")
        kind_name = value.get("kind")
        if not isinstance(name, str) or "\x00" in name or not isinstance(kind_name, str):
            raise RuntimeError("Backup contains an invalid registry value.")
        folded_name = name.casefold()
        if folded_name in value_names:
            raise RuntimeError("Backup contains duplicate registry values.")
        value_names.add(folded_name)
        _decode_registry_data(kind_name, value.get("data"))

    subkey_names: set[str] = set()
    for child in subkeys:
        if not isinstance(child, dict) or set(child) != {"name", "tree"}:
            raise RuntimeError("Backup contains an invalid registry subkey.")
        name = child.get("name")
        if not isinstance(name, str) or not name or "\x00" in name or "\\" in name or "/" in name:
            raise RuntimeError("Backup contains an invalid registry subkey name.")
        folded_name = name.casefold()
        if folded_name in subkey_names:
            raise RuntimeError("Backup contains duplicate registry subkeys.")
        subkey_names.add(folded_name)
        _validate_snapshot_tree(child.get("tree"), depth + 1)


def _restore_key_tree(winreg, root, path: str, tree: dict) -> None:
    """Restore a validated registry tree beneath one fixed HKCU path."""
    _, by_name = _registry_kind_maps(winreg)
    write = winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
    with winreg.CreateKeyEx(root, path, 0, write) as key:
        for value in tree["values"]:
            kind_name = value["kind"]
            try:
                kind = by_name[kind_name]
            except KeyError as exc:
                raise RuntimeError(
                    f"This Windows version does not support registry kind {kind_name}."
                ) from exc
            data = _decode_registry_data(kind_name, value["data"])
            winreg.SetValueEx(key, value["name"], 0, kind, data)
    for child in tree["subkeys"]:
        _restore_key_tree(winreg, root, path + "\\" + child["name"], child["tree"])


def _write_backup_manifest(target: str, operation: str, entries: list[dict]) -> None:
    """Atomically finalize and checksum a structured Explorer-view snapshot."""
    manifest = {
        "schema": BACKUP_SCHEMA_VERSION,
        "created_at": datetime.now(UTC).isoformat(),
        "operation": operation,
        "entries": entries,
    }
    manifest_path = os.path.join(target, BACKUP_MANIFEST)
    digest_path = os.path.join(target, BACKUP_DIGEST)
    pending_manifest = manifest_path + ".tmp"
    pending_digest = digest_path + ".tmp"
    payload = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if len(payload) > MAX_BACKUP_BYTES:
        raise RuntimeError("Backup manifest is too large to write safely.")
    digest = hashlib.sha256(payload).hexdigest()
    with open(pending_manifest, "xb") as handle:
        handle.write(payload)
        handle.flush()
        os.fsync(handle.fileno())
    with open(pending_digest, "x", encoding="ascii", newline="\n") as handle:
        handle.write(digest + "\n")
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(pending_manifest, manifest_path)
    os.replace(pending_digest, digest_path)
    try:
        os.remove(os.path.join(target, BACKUP_PENDING_MARKER))
    except FileNotFoundError:
        pass


def _backup_registry_state(operation: str = "unspecified") -> str:
    """Snapshot every affected HKCU key and return a backup directory.

    A manifest records both present and absent keys. Recovery first deletes all
    affected keys and then recreates only those that existed, which also removes
    keys created by a failed partial operation.
    """
    winreg = _open_winreg()
    target = _create_backup_directory()
    entries = []
    for key in BACKUP_KEYS:
        try:
            tree = _snapshot_key_tree(winreg, winreg.HKEY_CURRENT_USER, key)
        except FileNotFoundError:
            if _key_exists(winreg, key):
                raise
            LOG.info("Backup skipped for missing key HKCU\\%s", key)
            tree = None
        entries.append({"key": key, "existed": tree is not None, "tree": tree})

    _write_backup_manifest(target, operation, entries)
    return target


def _read_backup_manifest(backup: str) -> list[dict]:
    """Read and validate a Virelo Explorer-view backup manifest."""
    if os.path.exists(os.path.join(backup, BACKUP_PENDING_MARKER)):
        raise RuntimeError("Backup is incomplete and cannot be restored.")
    manifest_path = os.path.join(backup, BACKUP_MANIFEST)
    digest_path = os.path.join(backup, BACKUP_DIGEST)
    try:
        if os.path.getsize(manifest_path) > MAX_BACKUP_BYTES:
            raise RuntimeError("Backup manifest is too large to restore safely.")
        if os.path.getsize(digest_path) > 129:
            raise RuntimeError("Backup digest is too large to restore safely.")
        with open(manifest_path, "rb") as handle:
            payload = handle.read(MAX_BACKUP_BYTES + 1)
        with open(digest_path, encoding="ascii") as handle:
            expected_digest = handle.read(129).strip()
    except FileNotFoundError as exc:
        # Backups created before manifests cannot prove whether a missing .reg
        # file means the key was absent or an export failed partway through.
        # Treating that ambiguity as an absent key would delete live state
        # during restore, so automatic recovery must fail closed.
        raise RuntimeError("Legacy .reg backups cannot be restored automatically.") from exc
    except OSError as exc:
        raise RuntimeError(f"Backup manifest could not be read: {exc}") from exc

    if (
        len(expected_digest) != 64
        or any(character not in "0123456789abcdef" for character in expected_digest)
        or hashlib.sha256(payload).hexdigest() != expected_digest
    ):
        raise RuntimeError("Backup manifest failed its SHA-256 integrity check.")
    try:
        manifest = json.loads(payload.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError, RecursionError) as exc:
        raise RuntimeError(f"Backup manifest could not be read: {exc}") from exc

    if not isinstance(manifest, dict) or set(manifest) != {
        "schema",
        "created_at",
        "operation",
        "entries",
    }:
        raise RuntimeError("Backup manifest has an invalid structure.")
    if manifest.get("schema") != BACKUP_SCHEMA_VERSION:
        raise RuntimeError("Backup manifest uses an unsupported schema version.")
    if not isinstance(manifest.get("created_at"), str) or not isinstance(
        manifest.get("operation"), str
    ):
        raise RuntimeError("Backup manifest has invalid recovery metadata.")
    entries = manifest.get("entries")
    if not isinstance(entries, list) or len(entries) != len(BACKUP_KEYS):
        raise RuntimeError("Backup manifest does not describe every affected registry key.")

    by_key = {}
    for entry in entries:
        if not isinstance(entry, dict):
            raise RuntimeError("Backup manifest contains an unexpected registry key.")
        entry_key = entry.get("key")
        if not isinstance(entry_key, str) or entry_key not in BACKUP_KEYS:
            raise RuntimeError("Backup manifest contains an unexpected registry key.")
        if entry_key in by_key or not isinstance(entry.get("existed"), bool):
            raise RuntimeError("Backup manifest contains an invalid registry entry.")
        if set(entry) != {"key", "existed", "tree"}:
            raise RuntimeError("Backup manifest contains an invalid registry entry.")
        tree = entry.get("tree")
        if entry["existed"] != (tree is not None):
            raise RuntimeError("Backup manifest does not match its registry-key state.")
        if tree is not None:
            _validate_snapshot_tree(tree)
        by_key[entry_key] = entry
    if set(by_key) != set(BACKUP_KEYS):
        raise RuntimeError("Backup manifest is missing an affected registry key.")
    return [by_key[key] for key in BACKUP_KEYS]


def _restore_registry_state(backup: str) -> None:
    """Exactly restore all affected keys from a verified Virelo backup."""
    entries = _read_backup_manifest(backup)
    winreg = _open_winreg()

    for key in VIEW_CACHE_KEYS:
        _delete_key_tree(winreg, winreg.HKEY_CURRENT_USER, key)

    for entry in entries:
        if entry["existed"]:
            _restore_key_tree(
                winreg,
                winreg.HKEY_CURRENT_USER,
                entry["key"],
                entry["tree"],
            )


def _latest_backup_directory() -> str | None:
    """Return the newest complete Explorer-view backup, if one exists."""
    root = _backup_root()
    try:
        names = sorted(os.listdir(root), reverse=True)
    except OSError:
        return None
    for name in names:
        path = os.path.join(root, name)
        if not name.startswith("view-backup-") or not os.path.isdir(path):
            continue
        try:
            _read_backup_manifest(path)
            return path
        except RuntimeError:
            LOG.warning("Ignoring incomplete Explorer-view backup at %s", path)
    return None


def apply_details_default() -> dict:
    """Make Details view the default for all folders."""
    winreg = _open_winreg()
    backup = None
    try:
        backup = _backup_registry_state("apply-details")

        try:
            for key in VIEW_CACHE_KEYS:
                _delete_key_tree(winreg, winreg.HKEY_CURRENT_USER, key)

            _copy_key_tree(
                winreg,
                winreg.HKEY_LOCAL_MACHINE,
                FOLDER_TYPES_KEY,
                winreg.HKEY_CURRENT_USER,
                FOLDER_TYPES_KEY,
            )
            updated = _force_details_on_folder_types(winreg)
            _write_values(winreg, this_pc_bag_values())
        except Exception as exc:
            LOG.exception("Details-view registry update failed; restoring %s", backup)
            try:
                _restore_registry_state(backup)
            except Exception as rollback_exc:
                LOG.exception("Automatic Explorer-view rollback failed")
                return {
                    "ok": False,
                    "error": f"{exc}. Automatic rollback also failed: {rollback_exc}",
                    "data": {"backup": backup, "rolled_back": False},
                }
            return {
                "ok": False,
                "error": str(exc),
                "data": {"backup": backup, "rolled_back": True},
            }

        LOG.info(
            "Details view applied: %d folder views updated, backup at %s",
            updated,
            backup,
        )
        return {
            "ok": True,
            "data": {"updated": updated, "backup": backup, "restarted": False},
        }
    except Exception as e:
        LOG.exception("apply_details_default failed")
        data = {"backup": backup} if backup else {}
        return {"ok": False, "error": str(e), "data": data}


def reset_folder_views() -> dict:
    """Remove custom view state so Explorer returns to Windows defaults."""
    winreg = _open_winreg()
    backup = None
    try:
        backup = _backup_registry_state("reset-views")
        try:
            for key in VIEW_CACHE_KEYS:
                _delete_key_tree(winreg, winreg.HKEY_CURRENT_USER, key)
        except Exception as exc:
            LOG.exception("Folder-view reset failed; restoring %s", backup)
            try:
                _restore_registry_state(backup)
            except Exception as rollback_exc:
                LOG.exception("Automatic Explorer-view rollback failed")
                return {
                    "ok": False,
                    "error": f"{exc}. Automatic rollback also failed: {rollback_exc}",
                    "data": {"backup": backup, "rolled_back": False},
                }
            return {
                "ok": False,
                "error": str(exc),
                "data": {"backup": backup, "rolled_back": True},
            }
        LOG.info("Folder views reset, backup at %s", backup)
        return {"ok": True, "data": {"backup": backup, "restarted": False}}
    except Exception as e:
        LOG.exception("reset_folder_views failed")
        data = {"backup": backup} if backup else {}
        return {"ok": False, "error": str(e), "data": data}


def restore_latest_view_backup() -> dict:
    """Restore the latest complete Explorer-view backup."""
    target = _latest_backup_directory()
    if target is None:
        return {"ok": False, "error": "No complete Explorer view backup was found."}

    safety_backup = None
    try:
        safety_backup = _backup_registry_state("pre-restore")
        try:
            _restore_registry_state(target)
        except Exception as exc:
            LOG.exception("Restoring Explorer-view backup %s failed", target)
            try:
                _restore_registry_state(safety_backup)
            except Exception as rollback_exc:
                LOG.exception("Restoring the pre-restore safety backup failed")
                return {
                    "ok": False,
                    "error": f"{exc}. Safety rollback also failed: {rollback_exc}",
                    "data": {
                        "backup": target,
                        "safety_backup": safety_backup,
                        "rolled_back": False,
                    },
                }
            return {
                "ok": False,
                "error": str(exc),
                "data": {
                    "backup": target,
                    "safety_backup": safety_backup,
                    "rolled_back": True,
                },
            }
        LOG.info(
            "Explorer view backup restored: source=%s safety=%s",
            target,
            safety_backup,
        )
        return {
            "ok": True,
            "data": {
                "backup": target,
                "safety_backup": safety_backup,
                "restarted": False,
            },
        }
    except Exception as exc:
        LOG.exception("restore_latest_view_backup failed")
        data = {"backup": target}
        if safety_backup:
            data["safety_backup"] = safety_backup
        return {"ok": False, "error": str(exc), "data": data}
