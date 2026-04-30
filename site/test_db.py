import os
import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "analise_de_margem.settings")
django.setup()

from notas.models import OP

ops = OP.objects.prefetch_related('custo2_op_set').all()[:10]
for op in ops:
    has_history = op.custo2_op_set.exists()
    print(op.id_op, op.custo_2, has_history)
