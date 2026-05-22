from django.contrib import admin
from django.utils.html import format_html
from .models import FileConversion

@admin.action(description='Конвертировать выбранные файлы')
def convert_selected_files(modeladmin, request, queryset):
    for conversion in queryset.filter(status='pending'):
        conversion.convert_file()
    modeladmin.message_user(request, f"Запущена конвертация {queryset.count()} файлов")

class FileConversionAdmin(admin.ModelAdmin):
    list_display = ['id', 'original_file', 'target_format', 'status', 'created_at']
    list_filter = ['status', 'target_format']
    search_fields = ['original_file', 'id']
    readonly_fields = ['id', 'status', 'created_at', 'completed_at', 'error_message']
    actions = [convert_selected_files]
    
    fieldsets = (
        ('Основная информация', {
            'fields': ('id', 'original_file', 'target_format', 'status')
        }),
        ('Файлы', {
            'fields': ('converted_file',),
            'classes': ('collapse',)
        }),
        ('Статистика', {
            'fields': ('created_at', 'completed_at', 'error_message'),
            'classes': ('collapse',)
        }),
    )

# РЕГИСТРАЦИЯ МОДЕЛИ
admin.site.register(FileConversion, FileConversionAdmin)

# Кастомизация заголовков
admin.site.site_header = "Конвертер файлов - Администрирование"
admin.site.site_title = "Конвертер файлов"
admin.site.index_title = "Панель управления"
