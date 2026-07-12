FROM python:3.12-slim
WORKDIR /app
COPY admin_requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt
COPY admin_server.py /app/admin_server.py
COPY local_admin_server.py /app/local_admin_server.py
COPY release_store.py /app/release_store.py
COPY static /app/static
RUN mkdir -p /data/releases /data/license-data
EXPOSE 8080
CMD ["python", "/app/admin_server.py"]
