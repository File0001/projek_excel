"""Dataclass definition for File Task in the batch processing queue.

Extended with per-file configuration fields for batch conversion.
Adheres to camelCase variable naming and modular design.
"""

from dataclasses import dataclass, field
from typing import Literal, List, Optional


@dataclass
class FileTask:
    """Represents a single file conversion task in the queue.

    Attributes:
        id: Unique task identifier
        path: Absolute path to the file/folder/zip
        type: Task type - 'file', 'folder', or 'zip'
        status: Current status - 'pending', 'processing', 'completed', 'error'
        progress: Conversion progress percentage (0-100)
        parent_queue_id: Reference to queue that created this task (for ZIP extraction)
        
        Per-file configuration:
        selectedSheets: List of sheet names to convert (empty = all sheets)
        paperSize: Paper size - 'A4' or 'F4'
        orientation: Page orientation - 'Portrait' or 'Landscape'
        scaling: Scaling mode
        outputPath: Custom output path (None = same folder as source with auto-numbering)
    """
    id: str
    path: str
    type: Literal["file", "folder", "zip"]
    status: Literal["pending", "processing", "completed", "error"]
    progress: int = 0
    parent_queue_id: str = ""
    
    # Per-file configuration
    selectedSheets: List[str] = field(default_factory=list)
    paperSize: str = "A4"
    orientation: str = "Portrait"
    scaling: str = "Normal"
    outputPath: Optional[str] = None