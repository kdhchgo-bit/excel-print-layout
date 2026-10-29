"""Keep helper reports and renderings from accidentally overwriting inputs."""
from __future__ import annotations

import os
from pathlib import Path
import tempfile


def same_path(a: Path, b: Path) -> bool:
    if a.exists() and b.exists():
        return os.path.samefile(a, b)
    return os.path.normcase(str(a.expanduser().resolve())) == os.path.normcase(str(b.expanduser().resolve()))


def ensure_not_input(destination: Path, inputs: list[Path]) -> None:
    if any(same_path(destination, source) for source in inputs):
        raise ValueError("report path must not be an input workbook or PDF")


def write_bytes(destination: Path, data: bytes, overwrite: bool = False) -> str:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        if not destination.is_file():
            raise ValueError(f"output is not a file: {destination}")
        if destination.read_bytes() == data:
            return "unchanged"
        if not overwrite:
            raise FileExistsError(f"output exists with different contents: {destination}; choose a new path or --overwrite")
        with tempfile.NamedTemporaryFile("wb", dir=destination.parent, prefix=".report-", delete=False) as stream:
            temp = Path(stream.name)
            stream.write(data)
        try:
            os.replace(temp, destination)
        finally:
            if temp.exists():
                temp.unlink()
        return "replaced"
    with destination.open("xb") as stream:
        stream.write(data)
    return "created"
