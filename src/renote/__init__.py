"""Renote core package.

Modular utilities to transform wide PPTX slides into standard 16:9 format.

Modules:
- detectors: heuristics for slide classification (e.g., title slide detection)
- transforms: slide transformation strategies (title, split-into-thirds)
- processor: scenario handler/orchestrator
- utils: helper functions for XML cloning and layouts
"""

__all__ = [
    "detectors",
    "transforms",
    "processor",
    "utils",
]


