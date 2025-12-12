"""
Models for the accounts app.

This module contains:
- Custom User model with role field (ADMIN, OPERATOR, SUPERADMIN, USER)
- User-related models for authentication and authorization

Roles:
- ADMIN: Full access to Kartei, can edit all records, trigger sync
- OPERATOR: Limited access (SEPA restrictions, past-months limitations)
- SUPERADMIN: Approves/declines pending changes
- USER: Read-only access to history
"""

from __future__ import annotations

from django.contrib.auth.models import AbstractUser
from django.db import models


class UserRole(models.TextChoices):
    """
    User roles in the KindEltern system.
    
    Mirrors VBA role system from ExportProtection.GetCurrentUserRole:
    - ADMIN: "Admin" in Excel J1
    - OPERATOR: "Operator" in Excel J1
    - SUPERADMIN: Uses separate Suprime file
    - USER: View-only access
    """
    ADMIN = "ADMIN", "Admin"
    OPERATOR = "OPERATOR", "Operator"
    SUPERADMIN = "SUPERADMIN", "Superadmin"
    USER = "USER", "User"


class User(AbstractUser):
    """
    Custom User model with role field.
    
    Extends Django's AbstractUser to add:
    - role: The user's role in the system (ADMIN, OPERATOR, SUPERADMIN, USER)
    - Convenience properties for role checking
    """
    
    role = models.CharField(
        max_length=20,
        choices=UserRole.choices,
        default=UserRole.USER,
        help_text="User's role in the system. Determines access permissions.",
    )
    
    class Meta:
        db_table = "accounts_user"
        verbose_name = "User"
        verbose_name_plural = "Users"
    
    def __str__(self) -> str:
        return f"{self.username} ({self.get_role_display()})"
    
    # -------------------------------------------------------------------------
    # Role checking properties
    # -------------------------------------------------------------------------
    
    @property
    def is_admin_role(self) -> bool:
        """Check if user has Admin role."""
        return self.role == UserRole.ADMIN
    
    @property
    def is_operator(self) -> bool:
        """Check if user has Operator role."""
        return self.role == UserRole.OPERATOR
    
    @property
    def is_superadmin(self) -> bool:
        """Check if user has Superadmin role."""
        return self.role == UserRole.SUPERADMIN
    
    @property
    def is_user_role(self) -> bool:
        """Check if user has basic User role."""
        return self.role == UserRole.USER
    
    @property
    def can_edit_kartei(self) -> bool:
        """
        Check if user can edit Kartei records.
        
        Only Admin and Operator can edit (with restrictions for Operator).
        """
        return self.role in (UserRole.ADMIN, UserRole.OPERATOR)
    
    @property
    def can_approve_changes(self) -> bool:
        """
        Check if user can approve/decline pending changes.
        
        Only Superadmin can approve/decline.
        """
        return self.role == UserRole.SUPERADMIN
    
    @property
    def has_sepa_restrictions(self) -> bool:
        """
        Check if user has SEPA editing restrictions.
        
        Only Operator has SEPA restrictions (cannot edit SEPA rows).
        Mirrors VBA: ExportSyncKartei.ValidateOperatorSepaRestrictions
        """
        return self.role == UserRole.OPERATOR
    
    @property
    def has_past_months_restrictions(self) -> bool:
        """
        Check if user has past-months editing restrictions.
        
        Only Operator has past-months restrictions.
        Mirrors VBA: ExportProtection.ValidateAndFixPastMonths
        """
        return self.role == UserRole.OPERATOR
