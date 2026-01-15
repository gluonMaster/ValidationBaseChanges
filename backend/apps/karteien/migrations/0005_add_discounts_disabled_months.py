# Generated migration for discounts_disabled_months field
# Adds JSONField to store list of months (1-12) for which discounts are disabled

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('karteien', '0004_add_legacy_teacher_contract_fields'),
    ]

    operations = [
        migrations.AddField(
            model_name='karteirecord',
            name='discounts_disabled_months',
            field=models.JSONField(
                default=list,
                blank=True,
                verbose_name='Rabatte deaktiviert für Monate',
                help_text=(
                    'List of months (1-12) for which discounts are disabled. '
                    'Empty list = discounts disabled for ALL months (when discounts_disabled=True). '
                    'Non-empty list = discounts disabled only for these months.'
                ),
            ),
        ),
    ]
