from django.db import models
from django.contrib.auth.models import User

class Evaluacion(models.Model):
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name="evaluaciones")
    entidad = models.CharField(max_length=255, db_index=True)
    periodo = models.CharField(max_length=255, db_index=True)
    tipo = models.CharField(max_length=50, default="Capital")
    datos_cuestionario = models.TextField()
    fecha_creacion = models.CharField(max_length=50, blank=True, null=True)
    ultima_modificacion = models.CharField(max_length=50, blank=True, null=True)

    class Meta:
        verbose_name = "Evaluación"
        verbose_name_plural = "Evaluaciones"

    def __str__(self):
        return f"{self.entidad} - {self.periodo} ({self.tipo})"

