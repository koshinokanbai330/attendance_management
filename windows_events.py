"""
windows_events.py – Query Windows Event Log for system sleep / lock / shutdown times.

Used by AttendanceManager to auto-fill a missing end time for the previous day
when the user forgot to press the "終業" button.

Conditions (requirement 11):
  (a) Hibernate for 3+ hours  → use hibernate-start time as end time
  (b) Lock state for 3+ hours → use lock-start time as end time
  (c) Shutdown, 3+ hours ago  → use shutdown time as end time

On non-Windows platforms all functions return None silently.
"""

import subprocess
import sys
from datetime import date, datetime, timedelta
from typing import List, Optional, Tuple

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _run_powershell(script: str) -> List[Tuple[datetime, int]]:
    """Execute *script* in PowerShell and parse 'YYYY-MM-DD HH:MM:SS,<ID>' lines."""
    if sys.platform != "win32":
        return []
    try:
        result = subprocess.run(
            [
                "powershell",
                "-NonInteractive",
                "-NoProfile",
                "-Command",
                script,
            ],
            capture_output=True,
            text=True,
            timeout=30,
            encoding="utf-8",
        )
        events: List[Tuple[datetime, int]] = []
        for line in result.stdout.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split(",", 1)
            if len(parts) == 2:
                try:
                    dt = datetime.strptime(parts[0].strip(), "%Y-%m-%d %H:%M:%S")
                    eid = int(parts[1].strip())
                    events.append((dt, eid))
                except (ValueError, TypeError):
                    pass
        return events
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return []


def _query_log(log_name: str, event_ids: List[int], target_date: date) -> List[Tuple[datetime, int]]:
    """Return all events matching *event_ids* in *log_name* on *target_date* and the next day."""
    start = target_date.strftime("%Y-%m-%dT00:00:00")
    next_day = (target_date + timedelta(days=2)).strftime("%Y-%m-%dT00:00:00")
    id_filter = " -or ".join(f"$_.Id -eq {eid}" for eid in event_ids)
    script = f"""
try {{
    Get-WinEvent -FilterHashtable @{{
        LogName='{log_name}'
        StartTime='{start}'
        EndTime='{next_day}'
    }} -ErrorAction SilentlyContinue |
    Where-Object {{ {id_filter} }} |
    Sort-Object TimeCreated |
    ForEach-Object {{
        $_.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss') + ',' + $_.Id
    }}
}} catch {{ }}
"""
    return _run_powershell(script)


def _paired_event_start(
    start_events: List[Tuple[datetime, int]],
    end_events: List[Tuple[datetime, int]],
    min_gap_seconds: int = 3 * 3600,
) -> Optional[datetime]:
    """Return the start time of the first *start_event* whose gap to the next
    *end_event* is >= *min_gap_seconds*.  Also counts if no end event is found
    and the time since the start event is >= *min_gap_seconds*.
    """
    now = datetime.now()
    for start_dt, _ in start_events:
        next_end: Optional[datetime] = None
        for end_dt, _ in sorted(end_events):
            if end_dt > start_dt:
                next_end = end_dt
                break
        if next_end is None:
            gap = (now - start_dt).total_seconds()
        else:
            gap = (next_end - start_dt).total_seconds()
        if gap >= min_gap_seconds:
            return start_dt
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def get_last_work_end_time(target_date: date) -> Optional[datetime]:
    """Return the inferred work end time for *target_date* from Windows events.

    Evaluates the three conditions in requirement 11 and returns the *earliest*
    qualifying event time so we pick the first time the user stopped working.

    Returns None if no qualifying event is found or when not running on Windows.
    """
    if sys.platform != "win32":
        return None

    now = datetime.now()
    candidates: List[datetime] = []

    # ------------------------------------------------------------------
    # (a) Hibernate / Sleep  –  System log
    #     Event 42  = Kernel-Power "entering sleep" (modern Windows)
    #     Event 506 = Kernel-Power "entering sleep" (older)
    #     Event 1   = Kernel-Power "leaving sleep"
    #     Event 507 = Kernel-Power "leaving sleep" (older)
    # ------------------------------------------------------------------
    sleep_start_events = _query_log("System", [42, 506], target_date)
    sleep_end_events = _query_log("System", [1, 507], target_date)
    # Filter to only those on the target date
    sleep_start_events = [(dt, eid) for dt, eid in sleep_start_events
                          if dt.date() == target_date]
    candidate = _paired_event_start(sleep_start_events, sleep_end_events)
    if candidate:
        candidates.append(candidate)

    # ------------------------------------------------------------------
    # (b) Lock  –  Security log
    #     Event 4800 = workstation locked
    #     Event 4801 = workstation unlocked
    # ------------------------------------------------------------------
    lock_events = _query_log("Security", [4800], target_date)
    unlock_events = _query_log("Security", [4801], target_date)
    lock_events = [(dt, eid) for dt, eid in lock_events if dt.date() == target_date]
    candidate = _paired_event_start(lock_events, unlock_events)
    if candidate:
        candidates.append(candidate)

    # ------------------------------------------------------------------
    # (c) Shutdown  –  System log
    #     Event 6006 = EventLog service stopped (clean shutdown)
    #     Event 1074 = initiated shutdown / restart
    # ------------------------------------------------------------------
    shutdown_events = _query_log("System", [6006, 1074], target_date)
    for shutdown_dt, _ in shutdown_events:
        if shutdown_dt.date() == target_date:
            if (now - shutdown_dt).total_seconds() >= 3 * 3600:
                candidates.append(shutdown_dt)

    if not candidates:
        return None

    # Return the earliest qualifying time (most likely end-of-work)
    return min(candidates)
