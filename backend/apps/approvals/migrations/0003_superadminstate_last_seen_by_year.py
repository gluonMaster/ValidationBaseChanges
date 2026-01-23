# Generated manually for PROMPT_101: Add per-year last_seen_id tracking

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('approvals', '0002_add_admin_comment_to_pending_change'),
    ]

    operations = [
        migrations.AddField(
            model_name='superadminstate',
            name='last_seen_by_year',
            field=models.JSONField(
                blank=True,
                default=dict,
                help_text=(
                    "Mapping: year (as string) -> last_seen_id for NeuList. "
                    "Enables correct per-year tracking since record IDs are only unique "
                    "within a year (domain key is (year, id))."
                ),
            ),
        ),
    ]
