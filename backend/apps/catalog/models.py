"""
Catalog models for reference data: Teachers, Subjects, Teaching Assignments,
Price Options, Discounts, and Semester Configuration.

These tables provide standardized reference data for future forms and reporting,
without replacing the existing KarteiRecord.subject1/subject2 fields.

Discounts:
- Discount: справочник скидок (процентные и фиксированные)
- FamilyDiscount: скидки на семью в году
- RecordDiscount: скидки на конкретную запись

Semester Config:
- SemesterConfig: конфигурация границы семестра по годам
"""

import re
from decimal import Decimal

from django.conf import settings
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

    @property
    def badge_text(self) -> str:
        """
        Return formatted display text for badge:
        - PERCENT: "25%" (value * 100, rounded to integer)
        - FIXED: "10.00 €" (value as EUR amount)
        """
        if self.kind == DiscountKind.PERCENT:
            return f"{self.value * 100:.0f}%"
        else:
            return f"{self.value:.2f} €"

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


# =============================================================================
# Semester Configuration
# =============================================================================

class SemesterConfig(models.Model):
    """
    Per-year semester boundary configuration.

    Defines where semester 1 ends and semester 2 begins for each calendar year.
    If no row exists for a given year, the default boundary of 6 is used
    (semester 1 = months 1-6, semester 2 = months 7-12).
    """
    year = models.PositiveSmallIntegerField(
        unique=True,
        verbose_name='Jahr',
        help_text='Календарный год',
    )
    last_month_sem1 = models.PositiveSmallIntegerField(
        default=6,
        verbose_name='Letzter Monat Semester 1',
        help_text='Последний месяц семестра 1 (1-11)',
    )
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='Erstellt am')
    updated_at = models.DateTimeField(auto_now=True, verbose_name='Aktualisiert am')

    class Meta:
        verbose_name = 'Semester Config'
        verbose_name_plural = 'Semester Configs'
        ordering = ['-year']
        constraints = [
            models.CheckConstraint(
                check=models.Q(last_month_sem1__gte=1, last_month_sem1__lte=11),
                name='semester_boundary_range_1_to_11',
            ),
        ]

    def __str__(self):
        return f"{self.year}: Semester 1 → Monate 1-{self.last_month_sem1}"

    def clean(self):
        super().clean()
        # Prevent changing boundary for a year that already has KarteiRecords
        from apps.karteien.models import KarteiRecord

        if self.pk:
            # Editing existing config — check if year has records
            try:
                original = SemesterConfig.objects.get(pk=self.pk)
            except SemesterConfig.DoesNotExist:
                original = None

            if original and original.last_month_sem1 != self.last_month_sem1:
                if KarteiRecord.objects.filter(year=self.year).exists():
                    raise ValidationError({
                        'last_month_sem1': (
                            f'Die Semestergrenze für {self.year} kann nicht geändert werden, '
                            f'weil bereits Kartei-Einträge für dieses Jahr existieren.'
                        ),
                    })
        else:
            # New config — still check if trying to set non-default for year with records
            if KarteiRecord.objects.filter(year=self.year).exists():
                if self.last_month_sem1 != 6:
                    raise ValidationError({
                        'last_month_sem1': (
                            f'Die Semestergrenze für {self.year} kann nicht von 6 abweichen, '
                            f'weil bereits Kartei-Einträge für dieses Jahr existieren.'
                        ),
                    })

    @classmethod
    def get_boundary(cls, year: int) -> int:
        """Return last month of semester 1 for given year. Default: 6."""
        try:
            config = cls.objects.get(year=year)
            return config.last_month_sem1
        except cls.DoesNotExist:
            return 6


# =============================================================================
# Subject Category Models
# =============================================================================

class CategoryKind(models.TextChoices):
    """Type of subject category."""
    GROUP = 'GROUP', 'Gruppenunterricht'
    INDIVIDUAL = 'INDIVIDUAL', 'Einzelunterricht'


class SubjectCategory(models.Model):
    """
    Subject category with per-UE pricing rates.

    Categories group subjects by type (GROUP / INDIVIDUAL) for a given year
    and define the pricing rates per academic hour (45 min UE).
    """
    year = models.PositiveSmallIntegerField(
        verbose_name='Jahr',
        help_text='Календарный год',
    )
    name = models.CharField(
        max_length=200,
        verbose_name='Name',
        help_text='Название категории ("Русский", "Танцы" и т.д.)',
    )
    kind = models.CharField(
        max_length=20,
        choices=CategoryKind.choices,
        verbose_name='Art',
        help_text='Тип категории: групповой или индивидуальный',
    )
    yearly_rate = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        verbose_name='Jahresvertrag €/UE',
        help_text='Ставка €/UE (45 мин) для годичного контракта',
    )
    monthly_rate = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        verbose_name='Monatsvertrag €/UE',
        help_text='Ставка €/UE (45 мин) для месячного контракта',
    )
    group_threshold = models.PositiveSmallIntegerField(
        default=6,
        verbose_name='Schwellenwert Kleingruppe',
        help_text='Порог малокомплектности (только для GROUP)',
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Активна ли категория (мягкое удаление)',
    )

    class Meta:
        verbose_name = 'Subject Category'
        verbose_name_plural = 'Subject Categories'
        ordering = ['-year', 'name']
        constraints = [
            models.UniqueConstraint(
                fields=['year', 'name'],
                name='unique_category_year_name',
            ),
            models.CheckConstraint(
                check=models.Q(yearly_rate__gte=Decimal('0.00')),
                name='category_yearly_rate_non_negative',
            ),
            models.CheckConstraint(
                check=models.Q(monthly_rate__gte=Decimal('0.00')),
                name='category_monthly_rate_non_negative',
            ),
            models.CheckConstraint(
                check=models.Q(group_threshold__gte=1),
                name='category_group_threshold_min_1',
            ),
        ]
        indexes = [
            models.Index(fields=['year', 'kind'], name='idx_category_year_kind'),
            models.Index(fields=['year', 'is_active'], name='idx_category_year_active'),
        ]

    def __str__(self):
        return f"{self.year}: {self.name} ({self.kind})"

    def clean(self):
        super().clean()
        # For INDIVIDUAL categories, group_threshold is irrelevant — keep default
        # No extra validation needed; the field simply has no meaning for INDIVIDUAL.


class SubjectCategoryLink(models.Model):
    """
    Links a Subject to a SubjectCategory for a specific year.

    A subject may belong to at most one category per year
    (enforced by UniqueConstraint on subject + year).
    The ``year`` field is denormalized from ``category.year``
    and is automatically populated on save.
    """
    subject = models.ForeignKey(
        Subject,
        on_delete=models.PROTECT,
        related_name='category_links',
        verbose_name='Fach',
        help_text='Предмет',
    )
    category = models.ForeignKey(
        SubjectCategory,
        on_delete=models.CASCADE,
        related_name='links',
        verbose_name='Kategorie',
        help_text='Категория дисциплины',
    )
    year = models.PositiveSmallIntegerField(
        verbose_name='Jahr',
        help_text='Денормализовано из category.year',
        editable=False,
    )

    class Meta:
        verbose_name = 'Subject ↔ Category Link'
        verbose_name_plural = 'Subject ↔ Category Links'
        ordering = ['-year', 'category__name', 'subject__name']
        constraints = [
            models.UniqueConstraint(
                fields=['subject', 'year'],
                name='unique_subject_year_category',
            ),
        ]
        indexes = [
            models.Index(fields=['year', 'category'], name='idx_link_year_category'),
        ]

    def __str__(self):
        return f"{self.year}: {self.subject.name} → {self.category.name}"

    def clean(self):
        super().clean()
        if self.category_id:
            if self.year and self.year != self.category.year:
                raise ValidationError({
                    'year': 'Year must match category.year.',
                })

    def save(self, *args, **kwargs):
        if self.category_id:
            self.year = self.category.year
        super().save(*args, **kwargs)


# =============================================================================
# Discipline Group & Related Entry Models
# =============================================================================

class DisciplineGroup(models.Model):
    """
    One group per Subject per year.

    Created/activated when a Subject is linked to a GROUP category.
    Holds per-month duration entries and manual size overrides.
    """
    subject = models.ForeignKey(
        Subject,
        on_delete=models.PROTECT,
        related_name='discipline_groups',
        verbose_name='Fach',
        help_text='Предмет',
    )
    year = models.PositiveSmallIntegerField(
        verbose_name='Jahr',
        help_text='Учебный год',
    )
    category = models.ForeignKey(
        SubjectCategory,
        on_delete=models.PROTECT,
        related_name='groups',
        verbose_name='Kategorie',
        help_text='Категория (только GROUP)',
    )
    auto_scaling_enabled = models.BooleanField(
        default=False,
        verbose_name='Auto-Scaling',
        help_text='Включён ли автоскейлинг малокомплектных',
    )
    is_active = models.BooleanField(
        default=True,
        verbose_name='Aktiv',
        help_text='Мягкая деактивация группы',
    )

    class Meta:
        verbose_name = 'Discipline Group'
        verbose_name_plural = 'Discipline Groups'
        ordering = ['-year', 'subject__name']
        constraints = [
            models.UniqueConstraint(
                fields=['subject', 'year'],
                name='unique_discipline_group_subject_year',
            ),
        ]
        indexes = [
            models.Index(fields=['year', 'category'], name='idx_group_year_category'),
        ]

    def __str__(self):
        return f"{self.year}: {self.subject.name} ({self.category.name})"

    def clean(self):
        super().clean()
        if self.category_id:
            if self.category.kind != CategoryKind.GROUP:
                raise ValidationError({
                    'category': 'DisciplineGroup kann nur für GROUP-Kategorien erstellt werden.',
                })
            if self.category.year != self.year:
                raise ValidationError({
                    'category': 'Das Jahr der Kategorie muss mit dem Jahr der Gruppe übereinstimmen.',
                })


class DurationEntry(models.Model):
    """
    Per-month duration history for a DisciplineGroup.

    Each entry defines the lesson duration (in minutes) starting from
    ``effective_from_month`` and lasting until the next entry or end of year.
    """
    group = models.ForeignKey(
        DisciplineGroup,
        on_delete=models.CASCADE,
        related_name='duration_entries',
        verbose_name='Gruppe',
        help_text='Группа',
    )
    effective_from_month = models.PositiveSmallIntegerField(
        verbose_name='Ab Monat',
        help_text='Действует с месяца (1-12)',
    )
    duration_minutes = models.PositiveSmallIntegerField(
        verbose_name='Dauer (Min.)',
        help_text='Продолжительность в минутах',
    )
    changed_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        on_delete=models.SET_NULL,
        related_name='+',
        verbose_name='Geändert von',
        help_text='Кем изменено',
    )
    changed_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name='Geändert am',
        help_text='Когда изменено',
    )
    comment = models.TextField(
        blank=True,
        default='',
        verbose_name='Kommentar',
        help_text='Комментарий',
    )

    class Meta:
        verbose_name = 'Duration Entry'
        verbose_name_plural = 'Duration Entries'
        ordering = ['effective_from_month']
        constraints = [
            models.UniqueConstraint(
                fields=['group', 'effective_from_month'],
                name='unique_duration_group_month',
            ),
            models.CheckConstraint(
                check=models.Q(
                    effective_from_month__gte=1,
                    effective_from_month__lte=12,
                ),
                name='duration_month_range_1_12',
            ),
            models.CheckConstraint(
                check=models.Q(duration_minutes__gte=1),
                name='duration_minutes_min_1',
            ),
        ]

    def __str__(self):
        return f"Monat {self.effective_from_month}: {self.duration_minutes} Min."


class GroupSizeEntry(models.Model):
    """
    Manual group-size override for a DisciplineGroup.

    Each entry defines the manual size starting from ``effective_from_month``
    within one semester.  Overrides do **not** carry across semesters.
    """
    group = models.ForeignKey(
        DisciplineGroup,
        on_delete=models.CASCADE,
        related_name='size_entries',
        verbose_name='Gruppe',
        help_text='Группа',
    )
    effective_from_month = models.PositiveSmallIntegerField(
        verbose_name='Ab Monat',
        help_text='Действует с месяца (1-12)',
    )
    manual_size = models.PositiveSmallIntegerField(
        verbose_name='Manuelle Größe',
        help_text='Ручной размер группы',
    )
    changed_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        on_delete=models.SET_NULL,
        related_name='+',
        verbose_name='Geändert von',
        help_text='Кем изменено',
    )
    changed_at = models.DateTimeField(
        auto_now_add=True,
        verbose_name='Geändert am',
        help_text='Когда изменено',
    )
    comment = models.TextField(
        blank=True,
        default='',
        verbose_name='Kommentar',
        help_text='Комментарий',
    )

    class Meta:
        verbose_name = 'Group Size Entry'
        verbose_name_plural = 'Group Size Entries'
        ordering = ['effective_from_month']
        constraints = [
            models.UniqueConstraint(
                fields=['group', 'effective_from_month'],
                name='unique_size_group_month',
            ),
            models.CheckConstraint(
                check=models.Q(
                    effective_from_month__gte=1,
                    effective_from_month__lte=12,
                ),
                name='size_month_range_1_12',
            ),
            models.CheckConstraint(
                check=models.Q(manual_size__gte=1),
                name='manual_size_min_1',
            ),
        ]

    def __str__(self):
        return f"Monat {self.effective_from_month}: Größe {self.manual_size}"


# =============================================================================
# Discipline Group Helpers
# =============================================================================

def get_duration_for_month(
    group: DisciplineGroup, month: int, *, entries=None
) -> int | None:
    """
    Return duration_minutes for the given month.

    Uses the last DurationEntry whose ``effective_from_month <= month``.
    Returns ``None`` if no entry covers this month (duration not defined).

    Pass *entries* (pre-fetched queryset/list) to avoid an extra DB hit.
    """
    if entries is None:
        entries = group.duration_entries.all()

    best = None
    for entry in entries:
        if entry.effective_from_month <= month:
            if best is None or entry.effective_from_month > best.effective_from_month:
                best = entry
    return best.duration_minutes if best else None


def _semester_for_month(month: int, year: int) -> int:
    """Return 1 or 2 depending on SemesterConfig boundary for *year*."""
    boundary = SemesterConfig.get_boundary(year)
    return 1 if month <= boundary else 2


def get_manual_size_for_month(
    group: DisciplineGroup, month: int, year: int, *, entries=None
) -> int | None:
    """
    Return manual_size for the given month within the same semester.

    Uses the last GroupSizeEntry where:
      - ``effective_from_month <= month``
      - ``semester(effective_from_month, year) == semester(month, year)``

    Returns ``None`` if no manual override exists for this semester+month.

    Pass *entries* (pre-fetched queryset/list) to avoid an extra DB hit.
    """
    if entries is None:
        entries = group.size_entries.all()

    target_sem = _semester_for_month(month, year)

    best = None
    for entry in entries:
        if entry.effective_from_month <= month and _semester_for_month(entry.effective_from_month, year) == target_sem:
            if best is None or entry.effective_from_month > best.effective_from_month:
                best = entry
    return best.manual_size if best else None

