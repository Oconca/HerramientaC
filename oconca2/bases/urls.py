from django.urls import path
from . import views

app_name = "bases"

urlpatterns = [
    path('', views.index, name="index"),
    path('register', views.api_register, name="api_register"),
    path('token', views.api_token, name="api_token"),
    path('api/me', views.api_me, name="api_me"),
    path('api/guardar', views.api_guardar, name="api_guardar"),
    path('api/cargar', views.api_cargar, name="api_cargar"),
    path('api/evaluaciones', views.api_evaluaciones, name="api_evaluaciones"),
    path('api/borrar', views.api_borrar, name="api_borrar"),
    path('api/exportar', views.api_exportar, name="api_exportar"),
    path('api/exportar_word', views.api_exportar_word, name="api_exportar_word"),
    path('api/exportar_observaciones_word', views.api_exportar_observaciones_word, name="api_exportar_observaciones_word"),
    path('api/validar_exportacion', views.api_validar_exportacion, name="api_validar_exportacion"),
]