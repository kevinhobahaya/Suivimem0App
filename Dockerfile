# ------------ BASE PYTHON ------------
FROM python:3.11-slim


# ------------ ENVIRONNEMENT ------------
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1


# ------------ DEPENDANCES SYSTEME ------------
RUN apt-get update && apt-get install -y \
default-mysql-server \
default-mysql-client \
pkg-config \
build-essential \
libmysqlclient-dev \
&& rm -rf /var/lib/apt/lists/*


# ------------ DOSSIER DE L'APP ------------
WORKDIR /app


# ------------ COPIE DES FICHIERS ------------
COPY requirements.txt /app/


# ------------ INSTALLATION DES REQUIREMENTS ------------
RUN pip install --no-cache-dir -r requirements.txt


# ------------ COPIE DU PROJET ------------
COPY . /app/


# ------------ CONFIGURATION MYSQL ------------
RUN service mysql start && \
mysql -e "CREATE DATABASE IF NOT EXISTS ${DB_NAME:-mydb};" && \
mysql -e "ALTER USER 'root'@'localhost' IDENTIFIED WITH mysql_native_password BY '${DB_PASSWORD:-rootpass}';" && \
mysql -e "FLUSH PRIVILEGES;"


# ------------ EXPOSE LE PORT DJANGO ------------
EXPOSE 8000


# ------------ COMMANDES DE DEMARRAGE ------------
CMD service mysql start && \
python manage.py migrate && \
gunicorn nom_de_ton_projet.wsgi:application --bind 0.0.0.0:8000