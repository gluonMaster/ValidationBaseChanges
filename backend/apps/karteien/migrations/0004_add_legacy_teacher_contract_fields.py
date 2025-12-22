# Generated migration for adding legacy teacher names and contract type/status fields.
# See PROMPT_34_IMPORT_PATCH_CONTRACT_MARKERS_TEACHERS_AND_SEPA.md

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("karteien", "0003_add_billing_mode_fields"),
    ]

    operations = [
        # Legacy teacher names from Access Value11 / Value16
        migrations.AddField(
            model_name="karteirecord",
            name="teacher1_legacy_name",
            field=models.CharField(
                blank=True,
                default="",
                help_text="Teacher name for 1st semester (legacy text). Access: Value11.",
                max_length=255,
                verbose_name="Lehrer 1. HJ (Legacy)",
            ),
        ),
        migrations.AddField(
            model_name="karteirecord",
            name="teacher2_legacy_name",
            field=models.CharField(
                blank=True,
                default="",
                help_text="Teacher name for 2nd semester (legacy text). Access: Value16.",
                max_length=255,
                verbose_name="Lehrer 2. HJ (Legacy)",
            ),
        ),
        # Contract type raw and computed flag from Access Value14
        migrations.AddField(
            model_name="karteirecord",
            name="contract_type_raw",
            field=models.CharField(
                blank=True,
                default="",
                help_text="Raw contract type marker from Access Value14. May contain 'O/V' and other text.",
                max_length=255,
                verbose_name="Vertragstyp (Rohtext)",
            ),
        ),
        migrations.AddField(
            model_name="karteirecord",
            name="is_monthly_contract",
            field=models.BooleanField(
                default=False,
                help_text="True if contract_type_raw contains 'O/V' substring (case-insensitive).",
                verbose_name="Monatsvertrag (O/V)",
            ),
        ),
        # Contract status raw and computed flag from Access Value20
        migrations.AddField(
            model_name="karteirecord",
            name="contract_status_raw",
            field=models.CharField(
                blank=True,
                default="",
                help_text="Raw contract status marker from Access Value20. May contain 'KN' and other text.",
                max_length=255,
                verbose_name="Vertragsstatus (Rohtext)",
            ),
        ),
        migrations.AddField(
            model_name="karteirecord",
            name="is_contract_terminated",
            field=models.BooleanField(
                default=False,
                help_text="True if contract_status_raw contains 'KN' as separate token (word boundary).",
                verbose_name="Vertrag gekündigt (KN)",
            ),
        ),
    ]
