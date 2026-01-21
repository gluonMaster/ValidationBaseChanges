# Generated migration for adding months_csv_1 and months_csv_2 fields

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('karteien', '0007_add_end_month_fields'),
    ]

    operations = [
        migrations.AddField(
            model_name='karteirecord',
            name='months_csv_1',
            field=models.CharField(
                blank=True,
                default='',
                help_text='Optional. Comma-separated months for billing in 1st semester (1-6). Overrides start/end months.',
                max_length=50,
                verbose_name='Monate (CSV) 1. HJ',
            ),
        ),
        migrations.AddField(
            model_name='karteirecord',
            name='months_csv_2',
            field=models.CharField(
                blank=True,
                default='',
                help_text='Optional. Comma-separated months for billing in 2nd semester (7-12). Overrides start/end months.',
                max_length=50,
                verbose_name='Monate (CSV) 2. HJ',
            ),
        ),
    ]
