"""Queue Manager for batch processing Excel to PDF conversion tasks.

Handles queue persistence, task management, and sequential processing.
Adheres to camelCase variable naming and modular design.
"""

import json
import os
import sys
import uuid
from pathlib import Path
from typing import List, Dict, Any, Optional
from src.ui.file_task import FileTask
from src.utils.logger import appLogger


def _resolveQueueDataFile() -> Path:
    """Determine a writable location for the persisted queue data.

    In development the file lives next to this module. When frozen (installed
    build) the application folder is typically read-only, so we persist to
    %LOCALAPPDATA%\\CheapNFriendly instead.
    """
    if getattr(sys, "frozen", False):
        baseDir = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        dataDir = Path(baseDir) / "CheapNFriendly"
        try:
            dataDir.mkdir(parents=True, exist_ok=True)
        except Exception:
            dataDir = Path(os.path.expanduser("~")) / ".cheapnfriendly"
            dataDir.mkdir(parents=True, exist_ok=True)
        return dataDir / "queue_data.json"
    return Path(__file__).parent / "queue_data.json"


# Queue data file location
QUEUE_DATA_FILE = _resolveQueueDataFile()


class QueueManager:
    """Manages the batch processing queue with persistence support."""

    def __init__(self) -> None:
        """Initialize the queue manager and load existing queue from disk."""
        self.tasks: Dict[str, FileTask] = {}
        self._load_from_disk()

    def _generate_id(self) -> str:
        """Generate unique task ID."""
        return str(uuid.uuid4())[:8]

    def _save_to_disk(self) -> None:
        """Save queue state to disk as JSON."""
        try:
            data = {
                "tasks": {
                    task_id: {
                        "id": task.id,
                        "path": task.path,
                        "type": task.type,
                        "status": task.status,
                        "progress": task.progress,
                        "parent_queue_id": task.parent_queue_id,
                        # Per-file configuration
                        "selectedSheets": task.selectedSheets,
                        "paperSize": task.paperSize,
                        "orientation": task.orientation,
                        "scaling": task.scaling,
                        "outputPath": task.outputPath,
                    }
                    for task_id, task in self.tasks.items()
                }
            }
            with open(QUEUE_DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as save_err:
            appLogger.warning(f"Gagal menyimpan queue ke disk: {save_err}")

    def _load_from_disk(self) -> None:
        """Load queue state from disk JSON file."""
        if not QUEUE_DATA_FILE.exists():
            return

        try:
            with open(QUEUE_DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            for task_id, task_data in data.get("tasks", {}).items():
                self.tasks[task_id] = FileTask(
                    id=task_data["id"],
                    path=task_data["path"],
                    type=task_data["type"],
                    status=task_data["status"],
                    progress=task_data.get("progress", 0),
                    parent_queue_id=task_data.get("parent_queue_id", ""),
                    # Per-file configuration
                    selectedSheets=task_data.get("selectedSheets", []),
                    paperSize=task_data.get("paperSize", "A4"),
                    orientation=task_data.get("orientation", "Portrait"),
                    scaling=task_data.get("scaling", "Normal"),
                    outputPath=task_data.get("outputPath"),
                )
            appLogger.info(f"Loaded {len(self.tasks)} tasks from queue disk storage")
        except Exception as load_err:
            appLogger.warning(f"Gagal memuat queue dari disk: {load_err}")
            self.tasks = {}

    def add_file_task(self, file_path: str, task_type: str, parent_queue_id: str = "") -> FileTask:
        """Add a new file/folder/zip task to the queue."""
        task_id = self._generate_id()
        task = FileTask(
            id=task_id,
            path=file_path,
            type=task_type,
            status="pending",
            progress=0,
            parent_queue_id=parent_queue_id,
        )
        self.tasks[task_id] = task
        self._save_to_disk()
        appLogger.info(f"Added {task_type} task to queue: {file_path}")
        return task

    def remove_task(self, task_id: str) -> bool:
        """Remove a task from the queue."""
        if task_id in self.tasks:
            del self.tasks[task_id]
            self._save_to_disk()
            appLogger.info(f"Removed task from queue: {task_id}")
            return True
        return False

    def clear_all(self) -> None:
        """Clear all tasks from the queue."""
        self.tasks.clear()
        self._save_to_disk()
        appLogger.info("Cleared all tasks from queue")

    def get_task(self, task_id: str) -> Optional[FileTask]:
        """Get a specific task by ID."""
        return self.tasks.get(task_id)

    def get_all_tasks(self) -> List[FileTask]:
        """Get all tasks in the queue."""
        return list(self.tasks.values())

    def get_pending_tasks(self) -> List[FileTask]:
        """Get all pending tasks sorted by ID."""
        return sorted(
            [t for t in self.tasks.values() if t.status == "pending"],
            key=lambda x: x.id
        )

    def update_task_status(self, task_id: str, status: str, progress: int = 0) -> bool:
        """Update task status and progress."""
        if task_id in self.tasks:
            self.tasks[task_id].status = status
            self.tasks[task_id].progress = progress
            self._save_to_disk()
            return True
        return False

    def get_pending_count(self) -> int:
        """Get count of pending tasks."""
        return len([t for t in self.tasks.values() if t.status == "pending"])

    def get_processing_count(self) -> int:
        """Get count of tasks currently being processed."""
        return len([t for t in self.tasks.values() if t.status == "processing"])

    def update_task_config(
        self,
        task_id: str,
        config: Dict[str, Any],
        reactivate: bool = True,
    ) -> bool:
        """Update configuration for a specific task.

        Args:
            task_id: The task ID to update
            config: Dictionary with keys: selectedSheets, paperSize, orientation, scaling, outputPath
            reactivate: When True (default), a task that already finished (or
                failed) is reactivated back to 'pending' after its config changes,
                so it can be converted again without re-uploading the file.

        Returns:
            True if task was found and updated, False otherwise
        """
        if task_id in self.tasks:
            task = self.tasks[task_id]
            if "selectedSheets" in config:
                task.selectedSheets = config["selectedSheets"]
            if "paperSize" in config:
                task.paperSize = config["paperSize"]
            if "orientation" in config:
                task.orientation = config["orientation"]
            if "scaling" in config:
                task.scaling = config["scaling"]
            if "outputPath" in config:
                task.outputPath = config["outputPath"]

            # Reuse-after-conversion: when a finished (or failed) task has its
            # configuration changed, reactivate it so it becomes convertible
            # again without re-uploading the file.
            if reactivate and task.status in ("completed", "error"):
                task.status = "pending"
                task.progress = 0
                appLogger.info(
                    f"Task {task_id} reactivated to pending after config change"
                )

            self._save_to_disk()
            appLogger.info(f"Updated config for task {task_id}")
            return True
        return False

