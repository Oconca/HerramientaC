import os
from django.apps import AppConfig
from django.db.models.signals import post_migrate

def ensure_default_users(sender, **kwargs):
    from django.contrib.auth.models import User
    try:
        default_users = [
            {"username": "admin", "password": os.getenv("ADMIN_PASSWORD", "admin123"), "is_staff": True, "is_superuser": True},
            {"username": "usuario1", "password": "usuario123", "is_staff": False, "is_superuser": False},
            {"username": "usuario2", "password": "usuario123", "is_staff": False, "is_superuser": False},
        ]
        for u in default_users:
            user, created = User.objects.get_or_create(username=u["username"])
            if created or not user.has_usable_password():
                user.set_password(u["password"])
                user.is_staff = u["is_staff"]
                user.is_superuser = u["is_superuser"]
                user.save()
            elif u["is_superuser"] and (not user.is_staff or not user.is_superuser):
                user.is_staff = True
                user.is_superuser = True
                user.save()
    except Exception:
        pass

class BasesConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'bases'

    def ready(self):
        post_migrate.connect(ensure_default_users, sender=self)

