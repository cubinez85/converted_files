import os
import uuid
from django.conf import settings
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.permissions import IsAuthenticated
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework import status
from pathlib import Path

class ConvertFileView(APIView):
    permission_classes = [IsAuthenticated]
    parser_classes = [MultiPartParser, FormParser]

    def post(self, request, *args, **kwargs):
        file_obj = request.FILES.get('file')
        target_format = request.data.get('target_format')

        if not file_obj or not target_format:
            return Response({'error': 'Необходимо передать file и target_format'}, status=status.HTTP_400_BAD_REQUEST)

        # Валидация формата
        allowed = ['pdf', 'csv', 'excel', 'word', 'json']
        if target_format not in allowed:
            return Response({'error': f'Неподдерживаемый формат. Доступно: {allowed}'}, status=400)

        # Сохраняем во временную папку
        media_dir = Path(settings.MEDIA_ROOT) / 'conversions'
        media_dir.mkdir(parents=True, exist_ok=True)

        input_path = media_dir / f"{uuid.uuid4()}_{file_obj.name}"
        with open(input_path, 'wb+') as dest:
            for chunk in file_obj.chunks():
                dest.write(chunk)

        output_path = media_dir / f"{input_path.stem}.{target_format}"

        try:
            # 👇 Вызов вашей логики конвертации
            from converter import load_data, save_data  # импортируйте ваши функции
            data = load_data(str(input_path), target_format)
            save_data(data, target_format, str(output_path))

            # Удаляем исходник
            input_path.unlink(missing_ok=True)

            return Response({
                'status': 'success',
                'download_url': f'/media/conversions/{output_path.name}'
            }, status=status.HTTP_200_OK)

        except Exception as e:
            input_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)
            return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)
