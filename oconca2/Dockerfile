# Imagen base
FROM python:3.10.4-slim-bullseye

# Setear variables de entorno
ENV PIP_DISABLE_PIP_VERSION_CHECK 1
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Setear directorio de trabajo
WORKDIR /code

# Instalar dependencias
COPY oconca2/requirements.txt .
RUN pip install -r requirements.txt

# Copiar el proyecto
COPY oconca2/ .

# Exponer puerto por defecto
EXPOSE 8000

# Comando de arranque para Railway / Docker
CMD ["sh", "-c", "python manage.py collectstatic --noinput && python manage.py migrate && gunicorn app.wsgi:application --bind 0.0.0.0:${PORT:-8000} --workers 3 --timeout 120"]
