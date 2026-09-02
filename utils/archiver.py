"""
Сборщик zip-архивов в памяти.

Используется для выгрузки отчётов по цехам: один .xlsx на цех
внутри одного архива. Файлы накапливаются в словаре и упаковываются
одним вызовом build() / build_bytes().
"""

import io
import zipfile
from pathlib import PurePosixPath
from typing import Any, Optional


class ArchiveError(Exception):
    """Ошибка работы с архивом: пустое имя, дубликат файла, пустая сборка."""


class ReportArchiver:
    """
    Сборщик zip-архива из книг openpyxl и произвольных байтов.

    Файлы накапливаются в словаре и упаковываются одним вызовом build().
    Повторное имя без overwrite=True вызывает ArchiveError — совпадающие
    имена вызывающий код разводит суффиксом сам.
    """

    def __init__(self, compression: int = zipfile.ZIP_DEFLATED):
        """
        Создаёт пустой сборщик.

        Аргументы:
            compression: алгоритм сжатия zipfile. По умолчанию deflate;
                для архивов из чистых .xlsx (которые сами уже zip)
                ZIP_STORED быстрее при почти том же размере.
        """
        # {путь внутри архива: содержимое в байтах}
        self._files: dict[str, bytes] = {}
        self._compression = compression

    def _key(self, filename: str, folder: Optional[str]) -> str:
        """
        Собирает путь файла внутри архива из папки и имени.

        Путь нормализуется в posix-вид (слэши).

        Аргументы:
            filename: имя файла, непустое.
            folder: необязательная папка внутри архива.

        Возвращает:
            Нормализованный путь внутри архива.

        Исключения:
            ArchiveError: если имя файла пустое.
        """
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
        """
        Добавляет файл в будущий архив.

        Аргументы:
            filename: имя файла внутри архива.
            content: содержимое файла в байтах.
            folder: необязательная папка внутри архива.
            overwrite: разрешить перезапись существующего имени
                (по умолчанию повторное имя — ошибка).

        Возвращает:
            self — для цепочки вызовов.

        Исключения:
            ArchiveError: имя уже добавлено и overwrite=False.
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
        folder: Optional[str] = None,
        overwrite: bool = False,
    ) -> "ReportArchiver":
        """
        Сохраняет книгу openpyxl в байты и кладёт её в архив как .xlsx.

        Аргументы:
            workbook: объект openpyxl Workbook.
            filename: имя файла внутри архива.
            folder: необязательная папка внутри архива.
            overwrite: разрешить перезапись существующего имени.

        Возвращает:
            self — для цепочки вызовов.
        """
        buffer = io.BytesIO()
        workbook.save(buffer)
        return self.add_file(
            filename, buffer.getvalue(), folder=folder, overwrite=overwrite
        )

    @property
    def file_count(self) -> int:
        """
        Возвращает:
            Сколько файлов накоплено — для проверки «архив пустой».
        """
        return len(self._files)

    @property
    def filenames(self) -> list[str]:
        """
        Возвращает:
            Список путей файлов внутри архива — для диагностики и тестов.
        """
        return list(self._files)

    def build(self) -> io.BytesIO:
        """
        Упаковывает накопленные файлы в zip.

        Возвращает:
            Буфер, готовый к чтению (seek(0)).

        Исключения:
            ArchiveError: если не добавлено ни одного файла.
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
        """
        Собирает архив и отдаёт его содержимое как байты.

        Возвращает:
            Байты zip-архива — можно сразу класть в HttpResponse.
        """
        return self.build().getvalue()
