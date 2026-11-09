"""
Asynchronous Long-Running Job Manager (Inspired by Presenton FastMCP)

Enables long-running office tasks (mass slide generation, deep audits)
to execute in the background with immediate Job ID dispatch and polling,
preventing MCP connection timeouts.
"""

import time
import uuid
import threading
from typing import Dict, Any, Optional, Callable
from dataclasses import dataclass, field


@dataclass
class JobRecord:
    job_id: str
    command: str
    state: str = "pending"  # "pending" | "running" | "completed" | "failed"
    progress: int = 0  # 0 to 100
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)
    result: Optional[Any] = None
    error: Optional[str] = None


class AsyncJobManager:
    """In-memory asynchronous background task runner with status polling."""

    def __init__(self, ttl_seconds: int = 3600):
        self._jobs: Dict[str, JobRecord] = {}
        self._lock = threading.Lock()
        self._ttl_seconds = ttl_seconds

    def submit_job(self, command: str, worker_fn: Callable[..., Any], *args, **kwargs) -> str:
        job_id = f"job_{uuid.uuid4().hex[:12]}"
        record = JobRecord(job_id=job_id, command=command)

        with self._lock:
            self._jobs[job_id] = record

        def _runner():
            with self._lock:
                record.state = "running"
                record.progress = 10
                record.updated_at = time.time()

            try:
                # 실행
                result = worker_fn(*args, **kwargs)
                with self._lock:
                    record.state = "completed"
                    record.progress = 100
                    record.result = result
                    record.updated_at = time.time()
            except Exception as e:
                with self._lock:
                    record.state = "failed"
                    record.error = str(e)
                    record.updated_at = time.time()

        t = threading.Thread(target=_runner, daemon=True)
        t.start()
        return job_id

    def get_job_status(self, job_id: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            record = self._jobs.get(job_id)
            if not record:
                return None
            return {
                "job_id": record.job_id,
                "command": record.command,
                "state": record.state,
                "progress": record.progress,
                "elapsed_seconds": round(time.time() - record.created_at, 1),
                "result": record.result,
                "error": record.error
            }


# 전역 싱글턴 인스턴스
GLOBAL_JOB_MANAGER = AsyncJobManager()
