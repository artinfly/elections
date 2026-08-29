import io
import zipfile
from pathlib import PurePosixPath
from typing import Optional


class ArchiveError(Exception):
    """Ошибка работы с архивом: пустое имя, дубликат файла, пустая сборка."""

    pass


class ReportArchiver:
    """
    Сборщик zip-архива из книг openpyxl и произвольных байтов в памяти.
    Файлы накапливаются в словаре и упаковываются одним вызовом build().
    Используется для выгрузки отчётов по цехам (один .xlsx на цех).
    """

    def __init__(self, compression: int = zipfile.ZIP_DEFLATED):
        # {путь внутри архива: содержимое в байтах}
        self._files: dict[str, bytes] = {}
        # Алгоритм сжатия для zipfile (по умолчанию deflate)
        self._compression = compression

    def _key(self, filename: str, folder: Optional[str]) -> str:
        """
        Собирает путь файла внутри архива из папки и имени.
        Путь нормализуется в posix-вид (слэши), пустое имя — ошибка.
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
        Добавляет файл в будущий архив. Возвращает self для цепочки вызовов.
        Без overwrite=True повторное имя вызывает ArchiveError —
        вызывающий код сам разводит совпадающие имена суффиксом.
        """
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
        """
        Сохраняет книгу openpyxl в байты и кладёт её в архив как .xlsx-файл.
        """
        buffer = io.BytesIO()
        workbook.save(buffer)
        return self.add_file(
            filename, buffer.getvalue(), folder=folder, overwrite=overwrite
        )

    @property
    def file_count(self) -> int:
        """Сколько файлов накоплено в архиве (для проверки "архив пустой")."""
        return len(self._files)

    @property
    def filenames(self) -> list[str]:
        """Список путей файлов внутри архива (для диагностики и тестов)."""
        return list(self._files)

    def build(self) -> io.BytesIO:
        """
        Упаковывает накопленные файлы в zip и возвращает буфер,
        готовый к чтению (seek(0)). Пустой архив — ошибка ArchiveError.
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
        Собирает архив и отдаёт его содержимое как байты —
        удобно сразу положить в HttpResponse.
        """
        return self.build().getvalue()
