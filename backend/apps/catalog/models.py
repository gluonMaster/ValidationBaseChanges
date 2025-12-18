"""
Catalog models for reference data: Teachers, Subjects, Teaching Assignments, Price Options, and Discounts.

These tables provide standardized reference data for future forms and reporting,
without replacing the existing KarteiRecord.subject1/subject2 fields.

Discounts:
- Discount: справочник скидок (процентные и фиксированные)
- FamilyDiscount: скидки на семью в году
- RecordDiscount: скидки на конкретную запись
"""

import re
from decimal import Decimal

from django.core.exceptions import ValidationError
from django.core.validators import MinValueValidator, MaxValueValidator
from django.db import models


class Teacher(models.Model):
    """
    Teacher (Lehrer) reference table.
    
    Stores teacher information for linking to subjects via TeachingAssignment.
    """
    last_name = models.CharField(
        max_length=100,
        verbose_name='Nachname',
        help_text='Фамилия преподавателя'
    )
    first_name = models.CharField(
        max_length=100,
        verbose_name='Vorname',
        help_text='Имя преподавателя'
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Активен ли преподаватель'
    )

    class Meta:
        verbose_name = 'Teacher'
        verbose_name_plural = 'Teachers'
        ordering = ['last_name', 'first_name']
        constraints = [
            models.UniqueConstraint(
                fields=['last_name', 'first_name'],
                name='unique_teacher_name'
            )
        ]

    def __str__(self):
        return f"{self.last_name}, {self.first_name}"


class Subject(models.Model):
    """
    Subject (Fach) reference table.
    
    Stores standardized subject names for future forms and dropdowns.
    """
    name = models.CharField(
        max_length=200,
        unique=True,
        verbose_name='Name',
        help_text='Название предмета'
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Активен ли предмет'
    )

    class Meta:
        verbose_name = 'Subject'
        verbose_name_plural = 'Subjects'
        ordering = ['name']

    def __str__(self):
        return self.name


class TeachingAssignment(models.Model):
    """
    Teaching Assignment - links teachers to subjects for specific years.
    
    Allows queries:
    - By subject: get years and teachers
    - By teacher: get years and subjects  
    - By year: get list of subjects and who taught them
    """
    year = models.PositiveSmallIntegerField(
        db_index=True,
        verbose_name='Jahr',
        help_text='Учебный год (например, 2024, 2025)'
    )
    subject = models.ForeignKey(
        Subject,
        on_delete=models.PROTECT,
        related_name='assignments',
        verbose_name='Fach',
        help_text='Предмет'
    )
    teacher = models.ForeignKey(
        Teacher,
        on_delete=models.PROTECT,
        related_name='assignments',
        verbose_name='Lehrer',
        help_text='Преподаватель'
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Активно ли назначение'
    )

    class Meta:
        verbose_name = 'Teaching Assignment'
        verbose_name_plural = 'Teaching Assignments'
        ordering = ['-year', 'subject__name', 'teacher__last_name']
        constraints = [
            models.UniqueConstraint(
                fields=['year', 'subject', 'teacher'],
                name='unique_year_subject_teacher'
            )
        ]
        indexes = [
            models.Index(fields=['year', 'subject'], name='idx_year_subject'),
            models.Index(fields=['year', 'teacher'], name='idx_year_teacher'),
        ]

    def __str__(self):
        return f"{self.year}: {self.subject.name} — {self.teacher}"


# Regular expressions for detecting per-hour pricing subjects
INDIVIDUAL_PATTERN = re.compile(r'\bInd\.|VSpE_', re.IGNORECASE)
NACHHILFE_PATTERN = re.compile(r'\bNH\b|Nachhilfe', re.IGNORECASE)


class PriceOption(models.Model):
    """
    Price Option (Preisangabe) for a specific year and subject.
    
    Semantics of `amount`:
    - Default: price per month (€/Monat)
    - For subjects with "Ind." or "VSpE_" in the name (individual lessons): price per academic hour (€/UE)
    - For subjects with "NH" or "Nachhilfe" in the name (tutoring): price per academic hour (€/UE)
    """
    year = models.PositiveSmallIntegerField(
        db_index=True,
        verbose_name='Jahr',
        help_text='Учебный год (например, 2024, 2025)'
    )
    subject = models.ForeignKey(
        Subject,
        on_delete=models.PROTECT,
        related_name='price_options',
        verbose_name='Fach',
        help_text='Предмет'
    )
    amount = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        validators=[MinValueValidator(Decimal('0.00'))],
        verbose_name='Betrag',
        help_text='Сумма (€)'
    )
    comment = models.TextField(
        blank=True,
        default='',
        verbose_name='Kommentar',
        help_text='Комментарий (почему такая цена)'
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Активна ли эта цена'
    )

    class Meta:
        verbose_name = 'Price Option'
        verbose_name_plural = 'Price Options'
        ordering = ['-year', 'subject__name', 'amount']
        constraints = [
            models.UniqueConstraint(
                fields=['year', 'subject', 'amount', 'comment'],
                name='unique_year_subject_amount_comment'
            ),
            models.CheckConstraint(
                check=models.Q(amount__gte=Decimal('0.00')),
                name='price_amount_non_negative'
            )
        ]
        indexes = [
            models.Index(fields=['year', 'subject'], name='idx_price_year_subject'),
        ]

    def __str__(self):
        unit = self.get_price_unit()
        comment_suffix = f" ({self.comment[:30]}...)" if len(self.comment) > 30 else (f" ({self.comment})" if self.comment else "")
        return f"{self.year}: {self.subject.name} – {self.amount:.2f} {unit}{comment_suffix}"

    def get_price_unit(self) -> str:
        """
        Return the price unit based on subject name.
        
        Returns:
            '€/UE' for individual lessons (Ind., VSpE_) or tutoring (NH, Nachhilfe)
            '€/Monat' for regular subjects
        """
        subject_name = self.subject.name if self.subject_id else ""
        if INDIVIDUAL_PATTERN.search(subject_name) or NACHHILFE_PATTERN.search(subject_name):
            return '€/UE'
        return '€/Monat'

    def is_per_hour(self) -> bool:
        """Check if this price is per academic hour (UE) rather than per month."""
        return self.get_price_unit() == '€/UE'


# =============================================================================
# Discount Models
# =============================================================================

class DiscountKind(models.TextChoices):
    """Type of discount."""
    PERCENT = 'PERCENT', 'Prozent (%)'
    FIXED = 'FIXED', 'Fest (€)'


def validate_months_list(value):
    """
    Validate that months is a list of integers 1..12.
    """
    if value is None:
        return
    if not isinstance(value, list):
        raise ValidationError("Months must be a list.")
    if not value:
        raise ValidationError("Months list cannot be empty if provided.")
    for m in value:
        if not isinstance(m, int) or m < 1 or m > 12:
            raise ValidationError(f"Invalid month value: {m}. Must be 1-12.")
    if len(value) != len(set(value)):
        raise ValidationError("Months list contains duplicates.")


class Discount(models.Model):
    """
    Discount catalog entry.
    
    Types:
    - PERCENT: value is a fraction (0.00 - 0.99), e.g., 0.25 = 25%
    - FIXED: value is an amount in EUR, subtracted after percent discount
    """
    kind = models.CharField(
        max_length=10,
        choices=DiscountKind.choices,
        default=DiscountKind.PERCENT,
        verbose_name='Art',
        help_text='Тип скидки: процентная или фиксированная'
    )
    value = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        validators=[MinValueValidator(Decimal('0.00'))],
        verbose_name='Wert',
        help_text='Для PERCENT: 0.00-0.99 (например 0.25 = 25%). Для FIXED: сумма в EUR.'
    )
    description = models.TextField(
        blank=True,
        default='',
        verbose_name='Beschreibung',
        help_text='Описание скидки'
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Активна ли скидка'
    )

    class Meta:
        verbose_name = 'Discount'
        verbose_name_plural = 'Discounts'
        ordering = ['kind', '-value']

    def __str__(self):
        if self.kind == DiscountKind.PERCENT:
            percent_display = f"{self.value * 100:.0f}%"
            return f"{percent_display} – {self.description}" if self.description else percent_display
        else:
            return f"{self.value:.2f} € – {self.description}" if self.description else f"{self.value:.2f} €"

    def clean(self):
        super().clean()
        if self.kind == DiscountKind.PERCENT:
            if self.value < Decimal('0.00') or self.value > Decimal('0.99'):
                raise ValidationError({
                    'value': 'For PERCENT discount, value must be between 0.00 and 0.99.'
                })
        else:  # FIXED
            if self.value < Decimal('0.00'):
                raise ValidationError({
                    'value': 'For FIXED discount, value must be non-negative.'
                })


class BaseDiscountAssignment(models.Model):
    """
    Abstract base model for discount assignments with month range support.
    """
    discount = models.ForeignKey(
        Discount,
        on_delete=models.PROTECT,
        verbose_name='Rabatt',
        help_text='Применяемая скидка'
    )
    start_month = models.PositiveSmallIntegerField(
        default=1,
        validators=[MinValueValidator(1), MaxValueValidator(12)],
        verbose_name='Von Monat',
        help_text='Начальный месяц (1-12)'
    )
    end_month = models.PositiveSmallIntegerField(
        default=12,
        validators=[MinValueValidator(1), MaxValueValidator(12)],
        verbose_name='Bis Monat',
        help_text='Конечный месяц (1-12)'
    )
    months = models.JSONField(
        null=True,
        blank=True,
        validators=[validate_months_list],
        verbose_name='Monate (Liste)',
        help_text='Если задано [1,2,3], диапазон start_month..end_month игнорируется'
    )
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Erstellt am')
    updated_at = models.DateTimeField(auto_now=True, verbose_name='Aktualisiert am')

    class Meta:
        abstract = True

    def clean(self):
        super().clean()
        # Validate month range
        if self.start_month > self.end_month:
            raise ValidationError({
                'start_month': 'Start month cannot be greater than end month.'
            })

    def get_applicable_months(self) -> list[int]:
        """
        Return list of months where this discount applies.
        If `months` is set, use it; otherwise use start_month..end_month range.
        """
        if self.months:
            return sorted(self.months)
        return list(range(self.start_month, self.end_month + 1))

    def months_display(self) -> str:
        """Human-readable display of applicable months."""
        months = self.get_applicable_months()
        if len(months) == 12:
            return "ganzes Jahr"
        elif len(months) == 1:
            return f"Monat {months[0]}"
        elif months == list(range(months[0], months[-1] + 1)):
            # Continuous range
            return f"Monate {months[0]}-{months[-1]}"
        else:
            return f"Monate {', '.join(map(str, months))}"


class FamilyDiscount(BaseDiscountAssignment):
    """
    Discount applied to all records of a family for a specific year.
    
    Linked by family_id (string) rather than FK to allow flexibility.
    """
    year = models.PositiveSmallIntegerField(
        db_index=True,
        verbose_name='Jahr',
        help_text='Учебный год'
    )
    family_id = models.CharField(
        max_length=50,
        db_index=True,
        verbose_name='FamilyID',
        help_text='Идентификатор семьи (как в KarteiRecord.family_id)'
    )

    class Meta:
        verbose_name = 'Family Discount'
        verbose_name_plural = 'Family Discounts'
        ordering = ['-year', 'family_id', 'start_month']
        indexes = [
            models.Index(fields=['year', 'family_id'], name='idx_family_discount_year_fam'),
        ]

    def __str__(self):
        return f"{self.year}: {self.family_id} – {self.discount} ({self.months_display()})"


class RecordDiscount(BaseDiscountAssignment):
    """
    Discount applied to a specific KarteiRecord.
    """
    record = models.ForeignKey(
        'karteien.KarteiRecord',
        on_delete=models.CASCADE,
        related_name='record_discounts',
        verbose_name='Kartei-Eintrag',
        help_text='Запись, к которой применяется скидка'
    )

    class Meta:
        verbose_name = 'Record Discount'
        verbose_name_plural = 'Record Discounts'
        ordering = ['-record__year', 'record__pkid', 'start_month']

    def __str__(self):
        return f"#{self.record.pkid}: {self.discount} ({self.months_display()})"

