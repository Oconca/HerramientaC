from django.core.management.base import BaseCommand
from bases.models import Evaluacion
from bases.validator import validar_evaluacion_completa
import json

class Command(BaseCommand):
    help = 'Ejecuta el proceso interno de validación de exportación a Excel y Word para evaluaciones'

    def add_arguments(self, parser):
        parser.add_argument('--id', type=int, help='ID específico de la evaluación a validar')
        parser.add_argument('--tipo', type=str, help='Filtrar por tipo (Capital o Congreso)')
        parser.add_argument('--json', action='store_true', help='Imprimir salida en formato JSON')

    def handle(self, *args, **options):
        ev_id = options.get('id')
        tipo = options.get('tipo')
        output_json = options.get('json')

        qs = Evaluacion.objects.all().order_by('-id')
        if ev_id:
            qs = qs.filter(id=ev_id)
        if tipo:
            qs = qs.filter(tipo=tipo)

        if not qs.exists():
            self.stdout.write(self.style.WARNING("No se encontraron evaluaciones para validar."))
            return

        resultados = []
        total = qs.count()
        validas = 0
        con_errores = 0

        self.stdout.write(self.style.MIGRATE_HEADING(f"Iniciando auditoría interna de {total} evaluación(es)..."))

        for ev in qs:
            res = validar_evaluacion_completa(ev)
            res["id"] = ev.id
            resultados.append(res)

            if res["valido"]:
                validas += 1
                if not output_json:
                    self.stdout.write(self.style.SUCCESS(
                        f"[PASS] ID {ev.id} | {ev.entidad} - {ev.periodo} ({ev.tipo}) | Calificación: {res['resumen_excel']['calificacion_final']}%"
                    ))
            else:
                con_errores += 1
                if not output_json:
                    self.stdout.write(self.style.ERROR(
                        f"[FAIL] ID {ev.id} | {ev.entidad} - {ev.periodo} ({ev.tipo})"
                    ))
                    for err in res["errores"]:
                        self.stdout.write(self.style.ERROR(f"   -> {err}"))

            if res.get("advertencias") and not output_json:
                for adv in res["advertencias"]:
                    self.stdout.write(self.style.WARNING(f"   (Advertencia) -> {adv}"))

        if output_json:
            self.stdout.write(json.dumps(resultados, indent=2, ensure_ascii=False))
        else:
            self.stdout.write("=" * 60)
            self.stdout.write(f"Resumen de validación: Total={total} | Válidas={validas} | Con errores={con_errores}")
            if con_errores == 0:
                self.stdout.write(self.style.SUCCESS("✓ Todas las exportaciones validadas son 100% consistentes con fórmulas y datos."))
            else:
                self.stdout.write(self.style.ERROR(f"✗ Se detectaron {con_errores} evaluación(es) con inconsistencias."))
