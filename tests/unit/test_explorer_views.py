"""Tests for the Explorer default-view plan builders (pure logic, no registry)."""

import hashlib
import json
from datetime import datetime
from types import SimpleNamespace

import pytest

from virelo.services.explorer_views import (
    BACKUP_DIGEST,
    BACKUP_KEYS,
    BACKUP_MANIFEST,
    BAGMRU_KEY,
    BAGS_KEY,
    FOLDER_TYPES_KEY,
    ICON_SIZE_DETAILS,
    LOGICAL_VIEW_MODE_DETAILS,
    MODE_DETAILS,
    STREAMS_DEFAULTS_KEY,
    THIS_PC_GUID,
    VIEW_CACHE_KEYS,
    _create_backup_directory,
    _read_backup_manifest,
    _restore_registry_state,
    _write_backup_manifest,
    apply_details_default,
    backup_dir_name,
    this_pc_bag_values,
    top_view_values,
)


def test_details_constants_match_winsetview():
    """Details view is LogicalViewMode=1, Mode=4, IconSize=16 (WinSetView index 1)."""
    assert LOGICAL_VIEW_MODE_DETAILS == 1
    assert MODE_DETAILS == 4
    assert ICON_SIZE_DETAILS == 16


def test_view_cache_keys_cover_both_hive_locations():
    """The cache wipe must clear Bags/BagMRU in both registry locations plus Streams."""
    assert BAGS_KEY in VIEW_CACHE_KEYS
    assert BAGMRU_KEY in VIEW_CACHE_KEYS
    assert STREAMS_DEFAULTS_KEY in VIEW_CACHE_KEYS
    assert FOLDER_TYPES_KEY in VIEW_CACHE_KEYS
    legacy = [k for k in VIEW_CACHE_KEYS if k.startswith("Software\\Microsoft\\Windows\\Shell")]
    assert len(legacy) == 2


def test_top_view_values_force_details():
    """Each TopViews entry gets LogicalViewMode=Details and the Details icon size."""
    values = top_view_values(r"FolderTypes\{guid}\TopViews\{view}")
    by_name = {v.name: v for v in values}
    assert by_name["LogicalViewMode"].data == LOGICAL_VIEW_MODE_DETAILS
    assert by_name["LogicalViewMode"].kind == "dword"
    assert by_name["IconSize"].data == ICON_SIZE_DETAILS
    assert all(v.key == r"FolderTypes\{guid}\TopViews\{view}" for v in values)


def test_this_pc_bag_values_target_bag_one():
    """This PC gets NodeSlot 1 and Details bags under Bags\\1\\Shell and Bags\\1\\ComDlg."""
    values = this_pc_bag_values()
    keys = {v.key for v in values}
    assert rf"{BAGS_KEY}\1\Shell\{THIS_PC_GUID}" in keys
    assert rf"{BAGS_KEY}\1\ComDlg\{THIS_PC_GUID}" in keys
    node_slot = [v for v in values if v.name == "NodeSlot"]
    assert len(node_slot) == 1
    assert node_slot[0].data == 1
    assert node_slot[0].key == BAGMRU_KEY + r"\0"
    shell_bag = [
        v for v in values if v.key == rf"{BAGS_KEY}\1\Shell\{THIS_PC_GUID}" and v.name == "Mode"
    ]
    assert shell_bag[0].data == MODE_DETAILS


def test_this_pc_pidl_is_binary():
    """The BagMRU slot value 0 holds the This PC PIDL as raw bytes."""
    values = this_pc_bag_values()
    pidl = [v for v in values if v.key == BAGMRU_KEY and v.name == "0"]
    assert len(pidl) == 1
    assert pidl[0].kind == "binary"
    assert isinstance(pidl[0].data, bytes)
    assert pidl[0].data.startswith(bytes.fromhex("14001F50"))


def test_backup_dir_name_is_timestamped():
    """Backup directories sort chronologically and include microseconds."""
    name = backup_dir_name(datetime(2026, 7, 16, 13, 5, 9, 123456))
    assert name == "view-backup-20260716-130509-123456"


def test_create_backup_directory_handles_timestamp_collision(tmp_path, monkeypatch):
    """Two backups created at the same instant receive distinct directories."""
    monkeypatch.setattr("virelo.services.explorer_views._backup_root", lambda: str(tmp_path))
    now = datetime(2026, 7, 16, 13, 5, 9, 123456)

    first = _create_backup_directory(now)
    second = _create_backup_directory(now)

    assert first != second
    assert second.endswith("-001")


def test_backup_manifest_round_trip_tracks_present_and_absent_keys(tmp_path):
    """Recovery manifests preserve the exact existence state of every key."""
    empty_tree: dict[str, list[object]] = {"values": [], "subkeys": []}
    entries = [
        {
            "key": key,
            "existed": index == 0,
            "tree": empty_tree if index == 0 else None,
        }
        for index, key in enumerate(BACKUP_KEYS)
    ]
    _write_backup_manifest(str(tmp_path), "test", entries)

    assert _read_backup_manifest(str(tmp_path)) == entries


def test_legacy_backup_without_manifest_fails_closed(tmp_path):
    """An old partial export must not be mistaken for originally absent keys."""
    first_key = BACKUP_KEYS[0]
    filename = f"00-{first_key.rsplit(chr(92), 1)[-1]}.reg"
    (tmp_path / filename).write_text("Windows Registry Editor Version 5.00\n", encoding="utf-8")

    with pytest.raises(RuntimeError, match="Legacy .reg backups"):
        _read_backup_manifest(str(tmp_path))


def test_restore_rejects_a_tampered_snapshot_before_deleting_keys(tmp_path, monkeypatch):
    """Restore rejects a snapshot that attempts to escape the fixed HKCU keys."""
    empty_tree: dict[str, list[object]] = {"values": [], "subkeys": []}
    entries = [
        {
            "key": key,
            "existed": index == 0,
            "tree": empty_tree if index == 0 else None,
        }
        for index, key in enumerate(BACKUP_KEYS)
    ]
    _write_backup_manifest(str(tmp_path), "test", entries)
    manifest_path = tmp_path / BACKUP_MANIFEST
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["entries"][0]["key"] = r"Software\Virelo\Unexpected"
    payload = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode("utf-8")
    manifest_path.write_bytes(payload)
    (tmp_path / BACKUP_DIGEST).write_text(
        hashlib.sha256(payload).hexdigest() + "\n",
        encoding="ascii",
    )
    fake_winreg = SimpleNamespace(HKEY_CURRENT_USER=object())
    deleted = []
    monkeypatch.setattr("virelo.services.explorer_views._open_winreg", lambda: fake_winreg)
    monkeypatch.setattr(
        "virelo.services.explorer_views._delete_key_tree",
        lambda *_args: deleted.append(True),
    )

    with pytest.raises(RuntimeError, match="unexpected registry key"):
        _restore_registry_state(str(tmp_path))

    assert deleted == []


def test_restore_rejects_a_checksummed_unexpected_key_before_deleting_keys(tmp_path, monkeypatch):
    """A valid checksum cannot authorize a registry path outside the fixed allowlist."""
    entries = [{"key": key, "existed": False, "tree": None} for key in BACKUP_KEYS]
    _write_backup_manifest(str(tmp_path), "test", entries)
    manifest_path = tmp_path / BACKUP_MANIFEST
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["entries"][0]["key"] = r"Software\Virelo\Unexpected"
    payload = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode("utf-8")
    manifest_path.write_bytes(payload)
    (tmp_path / BACKUP_DIGEST).write_text(
        hashlib.sha256(payload).hexdigest() + "\n",
        encoding="ascii",
    )
    fake_winreg = SimpleNamespace(HKEY_CURRENT_USER=object())
    deleted = []
    monkeypatch.setattr("virelo.services.explorer_views._open_winreg", lambda: fake_winreg)
    monkeypatch.setattr(
        "virelo.services.explorer_views._delete_key_tree",
        lambda *_args: deleted.append(True),
    )

    with pytest.raises(RuntimeError, match="unexpected registry key"):
        _restore_registry_state(str(tmp_path))

    assert deleted == []


def test_structured_restore_writes_only_fixed_hkcu_paths(tmp_path, monkeypatch):
    """Automatic restore never executes backup content or selects another hive."""
    empty_tree: dict[str, list[object]] = {"values": [], "subkeys": []}
    entries = [
        {
            "key": key,
            "existed": index == 0,
            "tree": empty_tree if index == 0 else None,
        }
        for index, key in enumerate(BACKUP_KEYS)
    ]
    _write_backup_manifest(str(tmp_path), "test", entries)
    current_user = object()
    fake_winreg = SimpleNamespace(HKEY_CURRENT_USER=current_user)
    deleted = []
    restored: list[tuple[object, str, dict[str, list[object]]]] = []
    monkeypatch.setattr("virelo.services.explorer_views._open_winreg", lambda: fake_winreg)
    monkeypatch.setattr(
        "virelo.services.explorer_views._delete_key_tree",
        lambda _winreg, root, key: deleted.append((root, key)),
    )
    monkeypatch.setattr(
        "virelo.services.explorer_views._restore_key_tree",
        lambda _winreg, root, key, tree: restored.append((root, key, tree)),
    )

    _restore_registry_state(str(tmp_path))

    assert deleted == [(current_user, key) for key in BACKUP_KEYS]
    assert restored == [(current_user, BACKUP_KEYS[0], empty_tree)]


def test_apply_details_rolls_back_after_partial_registry_failure(monkeypatch):
    """A mutation failure restores the verified backup before returning."""
    fake_winreg = type(
        "FakeWinreg",
        (),
        {"HKEY_CURRENT_USER": object(), "HKEY_LOCAL_MACHINE": object()},
    )()
    restored: list[str] = []
    monkeypatch.setattr("virelo.services.explorer_views._open_winreg", lambda: fake_winreg)
    monkeypatch.setattr(
        "virelo.services.explorer_views._backup_registry_state",
        lambda operation: r"C:\backup",
    )
    monkeypatch.setattr("virelo.services.explorer_views._delete_key_tree", lambda *args: None)
    monkeypatch.setattr(
        "virelo.services.explorer_views._copy_key_tree",
        lambda *args: (_ for _ in ()).throw(OSError("Copy failed.")),
    )
    monkeypatch.setattr("virelo.services.explorer_views._restore_registry_state", restored.append)

    result = apply_details_default()

    assert result["ok"] is False
    assert result["data"] == {"backup": r"C:\backup", "rolled_back": True}
    assert restored == [r"C:\backup"]
