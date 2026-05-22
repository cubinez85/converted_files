import os
import uuid
from django.db import models
from django.utils import timezone
from pathlib import Path
from django.core.validators import FileExtensionValidator

class FileConversion(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Ожидает'),
        ('processing', 'Обрабатывается'),
        ('completed', 'Завершено'),
        ('failed', 'Ошибка'),
    ]
    
    FORMAT_CHOICES = [
        ('pdf', 'PDF'),
        ('csv', 'CSV'),
        ('excel', 'Excel (XLSX)'),
        ('word', 'Word (DOCX)'),
        ('json', 'JSON'),
    ]
    
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    original_file = models.FileField(
        upload_to='conversions/original/%Y/%m/%d/',
        validators=[FileExtensionValidator(allowed_extensions=['xlsx', 'xls', 'csv', 'json', 'docx', 'pdf'])],
        verbose_name='Исходный файл'
    )
    target_format = models.CharField(
        max_length=10,
        choices=FORMAT_CHOICES,
        verbose_name='Целевой формат'
    )
    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES,
        default='pending',
        verbose_name='Статус'
    )
    converted_file = models.FileField(
        upload_to='conversions/result/%Y/%m/%d/',
        blank=True,
        null=True,
        verbose_name='Конвертированный файл'
    )
    error_message = models.TextField(
        blank=True,
        null=True,
        verbose_name='Сообщение об ошибке'
    )
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Создан')
    completed_at = models.DateTimeField(null=True, blank=True, verbose_name='Завершен')
    file_size_original = models.BigIntegerField(default=0, verbose_name='Размер исходного (байт)')
    file_size_converted = models.BigIntegerField(default=0, null=True, blank=True, verbose_name='Размер результата (байт)')
    
    class Meta:
        verbose_name = 'Конвертация файла'
        verbose_name_plural = 'Конвертация файлов'
        ordering = ['-created_at']
    
    def __str__(self):
        return f"{self.original_file.name} -> {self.target_format} ({self.status})"
    
    def convert_file(self):
        """Выполняет конвертацию файла"""
        import sys
        import uuid
        from pathlib import Path
        from django.conf import settings
        from django.utils import timezone

        self.status = 'processing'
        self.save()

        try:
            # Добавляем корень проекта в sys.path для импорта converter.py
            project_root = Path(__file__).parent.parent.parent
            if str(project_root) not in sys.path:
                sys.path.insert(0, str(project_root))

            from converter import load_data, save_data

            original_path = self.original_file.path
            media_root = Path(settings.MEDIA_ROOT)
            output_dir = media_root / 'conversions' / 'temp'
            output_dir.mkdir(parents=True, exist_ok=True)

            # Базовое имя для временных файлов
            output_filename = f"{uuid.uuid4()}_{Path(self.original_file.name).stem}.{self.target_format}"
            output_path = output_dir / output_filename

            # 1. Загружаем данные
            data = load_data(original_path, self.target_format)

            # 2. Сохраняем и ПОЛУЧАЕМ СПИСОК созданных файлов
            # (save_data должен возвращать list[Path])
            saved_files = save_data(data, self.target_format, output_path)

            # 3. Проверяем, что файлы действительно созданы
            if not saved_files:
                raise FileNotFoundError("Функция save_data не вернула ни одного файла.")

            # Для CSV может быть создано несколько файлов (по таблицам), берем первый основной
            actual_file_path = saved_files[0]

            if not actual_file_path.exists():
                raise FileNotFoundError(f"Ожидаемый файл не найден на диске: {actual_file_path}")

            # 4. Сохраняем результат в модель Django
            with open(actual_file_path, 'rb') as f:
                # Сохраняем под реальным именем, которое сгенерировал скрипт
                self.converted_file.save(actual_file_path.name, f, save=False)

            self.status = 'completed'
            self.file_size_converted = self.converted_file.size
            self.completed_at = timezone.now()
            self.save()

            # 5. Очищаем все временные файлы из папки temp
            for temp_path in saved_files:
                if temp_path.exists():
                    temp_path.unlink()

            return True

        except Exception as e:
            import traceback
            self.status = 'failed'
            self.error_message = f"{str(e)}\n\n{traceback.format_exc()}"
            self.completed_at = timezone.now()
            self.save()
            return False
    
    def get_file_size_display_mb(self, field='original'):
        """Отображение размера в МБ"""
        if field == 'original':
            size = self.file_size_original or (self.original_file.size if self.original_file else 0)
        else:
            size = self.file_size_converted or (self.converted_file.size if self.converted_file else 0)
        return f"{size / (1024 * 1024):.2f} МБ"
