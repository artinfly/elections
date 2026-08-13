import io
import zipfile
from pathlib import PurePosixPath
from typing import Optional


class ArchiveError(Exception):
    pass


class ReportArchiver:
    def __init__(self, compression: int = zipfile.ZIP_DEFLATED):
        self._files: dict[str, bytes] = {}
        self._compression = compression

    def _key(self, filename: str, folder: Optional[str]) -> str:
        if not filename:
            raise ArchiveError("Имя файла не может быть пустым")
        path = PurePosixPath(folder) / filename if folder else PurePosixPath(filename)
        return path.as_posix()

    def add_file(
        self,
        filename: str,
        content: bytes,
        folder: Optional[str] = None,
        overwrite: bool = False,
    ) -> "ReportArchiver":
        key = self._key(filename, folder)
        if key in self._files and not overwrite:
            raise ArchiveError(f"Файл '{key}' уже добавлен в архив")
        self._files[key] = content
        return self

    def add_workbook(
        self,
        workbook,
        filename: str,
        folder: Optional[str] = None,
        overwrite: bool = False,
    ) -> "ReportArchiver":
        buffer = io.BytesIO()
        workbook.save(buffer)
        return self.add_file(
            filename, buffer.getvalue(), folder=folder, overwrite=overwrite
        )

    def remove_file(self, filename: str, folder: Optional[str] = None) -> None:
        self._files.pop(self._key(filename, folder), None)

    def clear(self) -> None:
        self._files.clear()

    @property
    def file_count(self) -> int:
        return len(self._files)

    @property
    def filenames(self) -> list[str]:
        return list(self._files)

    def build(self) -> io.BytesIO:
        if not self._files:
            raise ArchiveError("Нечего собирать в архив")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", self._compression) as archive:
            for path, content in self._files.items():
                archive.writestr(path, content)
        buffer.seek(0)
        return buffer

    def build_bytes(self) -> bytes:
        return self.build().getvalue()
