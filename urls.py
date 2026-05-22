"""
URL configuration for converted_files project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.2/topics/http/urls/
"""

from django.contrib import admin
from django.urls import path
from django.http import JsonResponse
from django.conf import settings
from django.conf.urls.static import static

# Простая функция-заглушка для корня сайта, чтобы проверить, что Django работает
def api_root(request):
    return JsonResponse({
        "message": "Django Converter API is running.",
        "docs": "/admin/",
        "status": "healthy"
    })

urlpatterns = [
    path('admin/', admin.site.urls),
    # Корневой адрес
    path('', api_root, name='api-root'),
    
    # Сюда мы будем добавлять пути для конвертации, например:
    # path('api/convert/', ConvertFileView.as_view(), name='convert'),
]

# Настройка раздачи медиа-файлов (скачанных/конвертированных файлов)
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
