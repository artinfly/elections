"""
Модуль сборки ZIP-архивов в оперативной памяти.

Описание:
    Предоставляет класс ReportArchiver для накопления файлов (книг Excel или
    произвольных байтов) и их последующей упаковки в один ZIP-архив без
    создания временных файлов на диске. Используется для выгрузки отчётов по цехам.
"""

import io
import zipfile
from pathlib import PurePosixPath
from typing import Any, Optional


class ArchiveError(Exception):
    """
    Ошибка работы с архивом.

    Описание:
        Исключение, возникающее при некорректных операциях с архивом:
        пустое имя файла, попытка добавить дубликат без разрешения перезаписи,
        попытка собрать пустой архив.
    """

    pass


class ReportArchiver:
    """
    Сборщик ZIP-архива из книг openpyxl и произвольных байтов.

    Описание:
        Файлы накапливаются во внутреннем словаре и упаковываются в архив
        одним вызовом метода build() или build_bytes().
        Повторное добавление файла с тем же именем без флага overwrite
        вызывает ошибку, чтобы предотвратить случайную потерю данных.
    """

    def __init__(self, compression: int = zipfile.ZIP_DEFLATED):
        """
        Создаёт пустой сборщик архивов.

        Описание:
            Инициализирует хранилище файлов и задает алгоритм сжатия.

        Аргументы:
            compression: алгоритм сжатия из модуля zipfile.
                По умолчанию используется ZIP_DEFLATED (сжатие).
                Для архивов, состоящих только из .xlsx (которые сами являются
                zip-архивами), можно использовать ZIP_STORED для ускорения сборки.

        Возвращает:
            Экземпляр класса.
        """
        # Словарь: {путь внутри архива: содержимое в байтах}.
        self._files: dict[str, bytes] = {}
        self._compression = compression

    def _key(self, filename: str, folder: Optional[str]) -> str:
        """
        Формирует нормализованный путь файла внутри архива.

        Описание:
            Объединяет имя папки и имя файла, используя прямые слеши (/),
            как того требует формат ZIP. Проверяет имя на пустоту.

        Аргументы:
            filename: имя файла (не может быть пустым).
            folder: необязательная папка внутри архива.

        Возвращает:
            str: нормализованный путь внутри архива.

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

        Описание:
            Сохраняет содержимое файла в память. Если файл с таким именем
            уже существует и перезапись не разрешена, выбрасывает ошибку.

        Аргументы:
            filename: имя файла внутри архива.
            content: содержимое файла в байтах.
            folder: необязательная папка внутри архива.
            overwrite: если True, разрешает перезапись существующего файла.

        Возвращает:
            ReportArchiver: текущий экземпляр (для цепочки вызовов).

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
        Добавляет книгу Excel (openpyxl) в архив.

        Описание:
            Сериализует объект книги в байты формата .xlsx и добавляет
            в архив как обычный файл.

        Аргументы:
            workbook: объект книги openpyxl Workbook.
            filename: имя файла внутри архива.
            folder: необязательная папка внутри архива.
            overwrite: если True, разрешает перезапись существующего файла.

        Возвращает:
            ReportArchiver: текущий экземпляр (для цепочки вызовов).
        """
        buffer = io.BytesIO()
        workbook.save(buffer)
        return self.add_file(
            filename, buffer.getvalue(), folder=folder, overwrite=overwrite
        )

    @property
    def file_count(self) -> int:
        """
        Возвращает количество файлов, добавленных в архив.

        Описание:
            Используется для проверки "архив пустой" перед сборкой.

        Возвращает:
            int: количество файлов.
        """
        return len(self._files)

    @property
    def filenames(self) -> list[str]:
        """
        Возвращает список путей файлов внутри архива.

        Описание:
            Полезно для диагностики, логирования и написания тестов.

        Возвращает:
            list: список строк с путями файлов.
        """
        return list(self._files)

    def build(self) -> io.BytesIO:
        """
        Упаковывает накопленные файлы в ZIP-архив.

        Описание:
            Создает байтовый буфер и записывает в него все файлы.
            Курсор буфера устанавливается в начало для последующего чтения.

        Возвращает:
            io.BytesIO: буфер с данными архива, готовый к чтению.

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
        Собирает архив и возвращает его содержимое в виде байтов.

        Описание:
            Удобный метод для получения готовых байтов, которые можно
            сразу передать в HTTP-ответ (например, в HttpResponse).

        Возвращает:
            bytes: байтовое представление ZIP-архива.
        """
        return self.build().getvalue()
