# Generated migration for adding end_month_1 and end_month_2 fields

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('karteien', '0006_add_contract_terminated_from_month'),
    ]

    operations = [
        migrations.AddField(
            model_name='karteirecord',
            name='end_month_1',
            field=models.PositiveSmallIntegerField(
                blank=True,
                help_text='Optional. Last month for billing in 1st semester (1-6). Months after this are 0.00.',
                null=True,
                verbose_name='Endmonat 1. HJ',
            ),
        ),
        migrations.AddField(
            model_name='karteirecord',
            name='end_month_2',
            field=models.PositiveSmallIntegerField(
                blank=True,
                help_text='Optional. Last month for billing in 2nd semester (7-12). Months after this are 0.00.',
                null=True,
                verbose_name='Endmonat 2. HJ',
            ),
        ),
    ]
