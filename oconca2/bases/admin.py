from django.contrib import admin
from .models import Evaluacion

admin.site.site_header = "OCONCA - Panel de Administración"
admin.site.site_title = "OCONCA CIMTRA Admin"
admin.site.index_title = "Gestión de Evaluaciones y Usuarios"

@admin.register(Evaluacion)
class EvaluacionAdmin(admin.ModelAdmin):
    list_display = ('entidad', 'periodo', 'tipo', 'usuario', 'ultima_modificacion')
    search_fields = ('entidad', 'periodo', 'tipo')
    list_filter = ('tipo', 'periodo')

