"""Сборка ZIP-архивов в памяти без временных файлов на диске."""

import io
import zipfile
from pathlib import PurePosixPath
from typing import Any, Self


class ArchiveError(Exception):
    """Некорректная операция с архивом: пустое имя, дубликат без overwrite, пустая сборка."""


class ReportArchiver:
    """Накапливает книги openpyxl и байты, упаковывает их в один ZIP-архив."""

    def __init__(self, compression: int = zipfile.ZIP_DEFLATED):
        """
        Аргументы:
            compression: алгоритм из zipfile. Для архивов только из .xlsx
                выгоднее ZIP_STORED — книги уже сжаты, повторное сжатие пусто.
        """
        self._files: dict[str, bytes] = {}
        self._compression = compression

    def _key(self, filename: str, folder: str | None) -> str:
        """
        Нормализует путь файла внутри архива (прямые слеши, как требует ZIP).

        Исключения:
            ArchiveError: имя файла пустое.
        """
        if not filename:
            raise ArchiveError("Имя файла не может быть пустым")
        path = PurePosixPath(folder) / filename if folder else PurePosixPath(filename)
        return path.as_posix()

    def add_file(
        self,
        filename: str,
        content: bytes,
        folder: str | None = None,
        overwrite: bool = False,
    ) -> Self:
        """
        Добавляет файл в архив; возвращает self для цепочки вызовов.

        Исключения:
            ArchiveError: имя уже занято, а overwrite=False.
        """
        key = self._key(filename, folder)
        if key in self._files and not overwrite:
            raise ArchiveError(f"Файл '{key}' уже добавлен в архив")
        self._files[key] = content
        return self

    def add_workbook(
        self,
        workbook: Any,
        filename: str,
        folder: str | None = None,
        overwrite: bool = False,
    ) -> Self:
        """Сериализует книгу openpyxl в .xlsx и добавляет как обычный файл."""
        buffer = io.BytesIO()
        workbook.save(buffer)
        return self.add_file(
            filename, buffer.getvalue(), folder=folder, overwrite=overwrite
        )

    @property
    def file_count(self) -> int:
        """Количество добавленных файлов: проверка «архив пуст» перед сборкой."""
        return len(self._files)

    @property
    def filenames(self) -> list[str]:
        """Пути файлов внутри архива."""
        return list(self._files)

    def build(self) -> io.BytesIO:
        """
        Упаковывает файлы в ZIP и возвращает буфер, готовый к чтению.

        Исключения:
            ArchiveError: не добавлено ни одного файла.
        """
        if not self._files:
            raise ArchiveError("Нечего собирать в архив")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", self._compression) as archive:
            for path, content in self._files.items():
                archive.writestr(path, content)
        buffer.seek(0)
        return buffer

    def build_bytes(self) -> bytes:
        """Собирает архив и возвращает байты для прямой отдачи в HttpResponse."""
        return self.build().getvalue()
