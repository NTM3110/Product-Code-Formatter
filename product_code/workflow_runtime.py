"""Session-scoped artifacts and idempotent background workflow jobs."""

from __future__ import annotations

import hashlib
import ctypes
import json
import os
import shutil
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path


TERMINAL_JOB_STATUSES = {"succeeded", "failed"}


class WorkflowFailure(Exception):
    def __init__(
        self,
        code,
        message,
        *,
        stage="",
        field="",
        details=None,
        retryable=True,
        operation_id="",
    ):
        super().__init__(message)
        self.code = str(code or "WORKFLOW_ERROR")
        self.message = str(message or "Workflow processing failed.")
        self.stage = str(stage or "")
        self.field = str(field or "")
        self.details = details if isinstance(details, dict) else {}
        self.retryable = bool(retryable)
        self.operation_id = str(operation_id or "")

    def as_dict(self):
        return {
            "code": self.code,
            "message": self.message,
            "stage": self.stage,
            "field": self.field,
            "details": self.details,
            "retryable": self.retryable,
            "operation_id": self.operation_id,
        }


def stable_signature(value):
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def file_sha256(path, chunk_size=1024 * 1024):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        while True:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


class WorkflowSessionStore:
    def __init__(self, root_dir, upload_dir):
        self.root_dir = Path(root_dir)
        self.upload_dir = Path(upload_dir)
        self._lock = threading.RLock()
        self.root_dir.mkdir(parents=True, exist_ok=True)
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        self._cleanup_stale_sessions()
        self.session_id = uuid.uuid4().hex
        self.session_dir = self.root_dir / self.session_id
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self._manifest = {
            "session_id": self.session_id,
            "pid": os.getpid(),
            "created_at": time.time(),
            "updated_at": time.time(),
            "artifacts": {},
        }
        self._write_manifest()

    @property
    def manifest_path(self):
        return self.session_dir / "manifest.json"

    def _safe_upload_path(self, saved_name):
        name = Path(str(saved_name or "")).name
        if not name:
            raise ValueError("Artifact saved name is empty.")
        return self.upload_dir / name

    def _write_manifest(self):
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self._manifest["updated_at"] = time.time()
        temporary = self.manifest_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(self._manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(self.manifest_path)

    def _cleanup_stale_sessions(self):
        for session_dir in self.root_dir.iterdir():
            if not session_dir.is_dir():
                continue
            manifest_path = session_dir / "manifest.json"
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except Exception:
                manifest = {}
            try:
                pid = int(manifest.get("pid") or 0)
            except (TypeError, ValueError):
                pid = 0
            if pid and pid != os.getpid() and process_is_running(pid):
                continue
            artifacts = manifest.get("artifacts") if isinstance(manifest.get("artifacts"), dict) else {}
            for artifact in artifacts.values():
                saved_name = artifact.get("saved_name") if isinstance(artifact, dict) else ""
                if saved_name:
                    try:
                        self._safe_upload_path(saved_name).unlink(missing_ok=True)
                    except OSError:
                        pass
            shutil.rmtree(session_dir, ignore_errors=True)

    def register_file(self, saved_name, *, kind, original_name="", signature="", metadata=None, supersede_kind=False):
        path = self._safe_upload_path(saved_name)
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(f"Artifact file does not exist: {path.name}")
        artifact_id = uuid.uuid4().hex
        artifact = {
            "artifact_id": artifact_id,
            "saved_name": path.name,
            "original_name": str(original_name or path.name),
            "kind": str(kind or "file"),
            "signature": str(signature or ""),
            "sha256": file_sha256(path),
            "size": path.stat().st_size,
            "valid": True,
            "created_at": time.time(),
            "metadata": metadata if isinstance(metadata, dict) else {},
        }
        with self._lock:
            if supersede_kind:
                for existing in self._manifest["artifacts"].values():
                    if existing.get("kind") == artifact["kind"]:
                        existing["valid"] = False
            self._manifest["artifacts"][artifact_id] = artifact
            self._write_manifest()
        return dict(artifact)

    def artifact_by_signature(self, kind, signature):
        if not signature:
            return None
        with self._lock:
            artifacts = list(self._manifest["artifacts"].values())
        for artifact in reversed(artifacts):
            if artifact.get("kind") != kind or artifact.get("signature") != signature or not artifact.get("valid"):
                continue
            if self._safe_upload_path(artifact.get("saved_name")).exists():
                return dict(artifact)
        return None

    def artifact(self, artifact_id):
        with self._lock:
            artifact = self._manifest["artifacts"].get(str(artifact_id or ""))
        return dict(artifact) if isinstance(artifact, dict) else None

    def invalidate(self, *, kinds=None):
        allowed = set(kinds or [])
        changed = False
        with self._lock:
            for artifact in self._manifest["artifacts"].values():
                if allowed and artifact.get("kind") not in allowed:
                    continue
                if artifact.get("valid"):
                    artifact["valid"] = False
                    changed = True
            if changed:
                self._write_manifest()
        return changed

    def snapshot(self):
        with self._lock:
            return json.loads(json.dumps(self._manifest, ensure_ascii=False))

    def close(self):
        with self._lock:
            artifacts = list(self._manifest.get("artifacts", {}).values())
        for artifact in artifacts:
            saved_name = artifact.get("saved_name") if isinstance(artifact, dict) else ""
            if saved_name:
                try:
                    self._safe_upload_path(saved_name).unlink(missing_ok=True)
                except OSError:
                    pass
        shutil.rmtree(self.session_dir, ignore_errors=True)


def process_is_running(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if os.name == "nt":
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        process_query_limited_information = 0x1000
        from ctypes import wintypes
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        kernel32.GetExitCodeProcess.restype = wintypes.BOOL
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL
        still_active = 259
        handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
        if not handle:
            return ctypes.get_last_error() == 5
        try:
            exit_code = ctypes.c_ulong()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                return True
            return exit_code.value == still_active
        finally:
            kernel32.CloseHandle(handle)
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


class WorkflowJobManager:
    def __init__(self, max_workers=2):
        self._executor = ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="workflow-job")
        self._lock = threading.RLock()
        self._jobs = {}
        self._signature_jobs = {}

    def _public(self, job):
        return json.loads(json.dumps(job, ensure_ascii=False, default=str))

    def start(self, kind, signature, runner, *, retry=False, context=None):
        signature_key = f"{kind}:{signature}"
        with self._lock:
            existing_id = self._signature_jobs.get(signature_key)
            existing = self._jobs.get(existing_id) if existing_id else None
            if existing and existing.get("status") in {"queued", "running", "succeeded"}:
                return self._public(existing)
            if existing and existing.get("status") == "failed" and not retry:
                return self._public(existing)
            job_id = uuid.uuid4().hex
            job = {
                "job_id": job_id,
                "operation_id": job_id,
                "kind": str(kind),
                "signature": str(signature),
                "status": "queued",
                "progress": {"done": 0, "total": 1, "percent": 0, "label": "Đang chờ xử lý"},
                "context": context if isinstance(context, dict) else {},
                "result": None,
                "error": None,
                "created_at": time.time(),
                "updated_at": time.time(),
            }
            self._jobs[job_id] = job
            self._signature_jobs[signature_key] = job_id
            self._executor.submit(self._run, job_id, runner)
            return self._public(job)

    def _run(self, job_id, runner):
        self._update(job_id, status="running", label="Đang xử lý")

        def progress(done, total, label):
            self._update(job_id, done=done, total=total, label=label)

        try:
            result = runner(progress)
        except WorkflowFailure as exc:
            if not exc.operation_id:
                exc.operation_id = job_id
            self._update(job_id, status="failed", error=exc.as_dict(), label=exc.message)
        except Exception as exc:
            failure = WorkflowFailure(
                "UNEXPECTED_PROCESSING_ERROR",
                str(exc),
                retryable=True,
                operation_id=job_id,
            )
            self._update(job_id, status="failed", error=failure.as_dict(), label=failure.message)
        else:
            self._update(job_id, status="succeeded", result=result, done=1, total=1, label="Đã hoàn tất")

    def _update(self, job_id, *, status=None, result=None, error=None, done=None, total=None, label=None):
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            if status:
                job["status"] = status
            if result is not None:
                job["result"] = result
            if error is not None:
                job["error"] = error
            progress = job["progress"]
            safe_total = max(1, int(total if total is not None else progress.get("total", 1)))
            safe_done = max(0, min(int(done if done is not None else progress.get("done", 0)), safe_total))
            progress.update({
                "done": safe_done,
                "total": safe_total,
                "percent": int(round((safe_done / safe_total) * 100)),
                "label": str(label if label is not None else progress.get("label", "")),
            })
            job["updated_at"] = time.time()

    def get(self, job_id):
        with self._lock:
            job = self._jobs.get(str(job_id or ""))
            return self._public(job) if job else None

    def shutdown(self):
        self._executor.shutdown(wait=False, cancel_futures=True)
