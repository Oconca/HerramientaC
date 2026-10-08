import os
import sys
import django

os.environ['DJANGO_SETTINGS_MODULE'] = 'app.settings'
django.setup()

from django.core.management import call_command

print("=== CHECK DJANGO SETTINGS ===")
call_command('check')

print("\n=== COLLECTSTATIC DRY-RUN ===")
call_command('collectstatic', interactive=False, dry_run=True)

print("\n=== RUN TESTS BASES ===")
call_command('test', 'bases', keepdb=True, interactive=False)

print("\n=== RUN VALIDAR_EXPORTACIONES ===")
call_command('validar_exportaciones')

print("\nALL VERIFICATIONS PASSED SUCCESSFULLY!")
