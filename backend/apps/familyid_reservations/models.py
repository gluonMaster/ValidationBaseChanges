"""
FamilyID Reservation model.

Allows admins to reserve the next available FamilyID before actually creating a family.
This is useful when an admin knows they need to create a family but cannot do it immediately.

FamilyID format: "1. <number>" (global across all years).
"""

from django.conf import settings
from django.db import models


class FamilyIdReservation(models.Model):
    """
    Reservation of a FamilyID.
    
    When an admin wants to reserve the next available FamilyID without creating
    a family immediately, they can create a reservation. Later, when creating
    the family, they can use this reserved ID.
    """
    family_id = models.CharField(
        max_length=50,
        unique=True,
        db_index=True,
        verbose_name='Family ID',
        help_text='Reserved FamilyID in format "1. <number>"'
    )
    
    reserved_at = models.DateTimeField(
        auto_now_add=True,
        db_index=True,
        verbose_name='Reserved At',
        help_text='When the reservation was created'
    )
    
    reserved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name='familyid_reservations',
        verbose_name='Reserved By',
        help_text='User who created the reservation'
    )
    
    is_used = models.BooleanField(
        default=False,
        db_index=True,
        verbose_name='Is Used',
        help_text='Whether this reservation has been used to create a family'
    )
    
    used_at = models.DateTimeField(
        null=True,
        blank=True,
        verbose_name='Used At',
        help_text='When the reservation was used to create a family'
    )
    
    note = models.CharField(
        max_length=200,
        blank=True,
        default='',
        verbose_name='Note',
        help_text='Optional note about this reservation'
    )

    class Meta:
        verbose_name = 'FamilyID Reservation'
        verbose_name_plural = 'FamilyID Reservations'
        ordering = ['-reserved_at']
        indexes = [
            models.Index(
                fields=['is_used', 'reserved_at'],
                name='familyid_res_active_idx'
            ),
        ]

    def __str__(self):
        status = '(used)' if self.is_used else '(active)'
        return f"{self.family_id} {status}"
