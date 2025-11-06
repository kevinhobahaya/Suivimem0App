# Utiliser l'image officielle Python
FROM python:3.11-slim

# Définir le répertoire de travail
WORKDIR /app

# Copier les fichiers requirements.txt et installer les dépendances
COPY requirements.txt .
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Copier tout le code
COPY . .

# Exposer le port 8000 pour Django
EXPOSE 8000

# Commande pour lancer Django avec Gunicorn
CMD ["gunicorn", "SuiviMemo.wsgi:application", "--bind", "0.0.0.0:8000"]
