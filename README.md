# How to create jango
plz run command
```
docker-compose up -d --build
docker exec -it ed_app bash

# create DB
python manage.py makemigrations
python manage.py migrate

# create superuser
python manage.py createsuperuser
ユーザー名: [custom username]
メールアドレス: [custom e-mail]
Password: [custom passwd]
```

# Engineering_Design
