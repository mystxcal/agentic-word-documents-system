from __future__ import annotations

import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Iterator


@dataclass
class StageTimings:
    """Monotonic build-stage measurements suitable for human and JSON reports."""

    started_ns: int = field(default_factory=time.perf_counter_ns)
    stages: list[dict] = field(default_factory=list)

    @contextmanager
    def measure(self, name: str, **details) -> Iterator[None]:
        start = time.perf_counter_ns()
        outcome = "succeeded"
        try:
            yield
        except Exception:
            outcome = "failed"
            raise
        finally:
            elapsed_ms = (time.perf_counter_ns() - start) / 1_000_000
            self.stages.append(
                {
                    "name": name,
                    "duration_ms": round(elapsed_ms, 3),
                    "outcome": outcome,
                    **details,
                }
            )

    def snapshot(self) -> dict:
        total_ms = (time.perf_counter_ns() - self.started_ns) / 1_000_000
        grouped: dict[str, float] = {}
        for stage in self.stages:
            grouped[stage["name"]] = round(
                grouped.get(stage["name"], 0.0) + float(stage["duration_ms"]),
                3,
            )
        return {
            "total_ms": round(total_ms, 3),
            "stages": list(self.stages),
            "by_name_ms": grouped,
        }
