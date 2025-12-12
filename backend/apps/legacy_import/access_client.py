"""
Access database client for legacy_import.

This module provides functions for connecting to and reading data from
Access databases (.accdb files). It uses pyodbc with the Microsoft ACE
ODBC driver.

Configuration:
- ACCESS_CONN_STRING_TEMPLATE: Connection string template (env variable)
- ACCESS_BASE_DIR: Base directory for .accdb files (env variable)

Example connection string template:
    DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ={file_path};

For Windows: Uses Microsoft ACE ODBC driver
For Linux: Requires unixODBC + mdbtools (limited support)

See ARCHITECTURE.md section 2.7 for more details.
"""

from __future__ import annotations

import logging
import os
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Iterator

from django.conf import settings

if TYPE_CHECKING:
    pass

logger = logging.getLogger(__name__)


# =============================================================================
# Configuration
# =============================================================================

def get_access_base_dir() -> Path:
    """
    Get the base directory for Access database files.
    
    Returns:
        Path to the directory containing .accdb files.
        
    Raises:
        ValueError: If ACCESS_BASE_DIR is not configured.
    """
    base_dir = getattr(settings, "ACCESS_BASE_DIR", None) or os.environ.get("ACCESS_BASE_DIR")
    if not base_dir:
        raise ValueError(
            "ACCESS_BASE_DIR is not configured. Set it in settings.py or as an environment variable."
        )
    return Path(base_dir)


def get_connection_string_template() -> str:
    """
    Get the ODBC connection string template.
    
    The template should contain {file_path} placeholder which will be
    replaced with the actual .accdb file path.
    
    Returns:
        Connection string template.
    """
    default_template = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ={file_path};"
    return (
        getattr(settings, "ACCESS_CONN_STRING_TEMPLATE", None)
        or os.environ.get("ACCESS_CONN_STRING_TEMPLATE")
        or default_template
    )


def build_connection_string(file_path: Path) -> str:
    """
    Build an ODBC connection string for a specific Access file.
    
    Args:
        file_path: Path to the .accdb file.
        
    Returns:
        Complete connection string for pyodbc.
    """
    template = get_connection_string_template()
    return template.format(file_path=str(file_path))


def resolve_access_file(file_name_or_path: str) -> Path:
    """
    Resolve an Access file name or path to an absolute path.
    
    If the input is a relative path or just a file name, it is resolved
    relative to ACCESS_BASE_DIR. If it's an absolute path, it's used as-is.
    
    Args:
        file_name_or_path: File name (e.g., "KindElternDaten_25_front.accdb")
                          or full path.
                          
    Returns:
        Absolute path to the .accdb file.
        
    Raises:
        FileNotFoundError: If the file does not exist.
    """
    path = Path(file_name_or_path)
    
    if not path.is_absolute():
        base_dir = get_access_base_dir()
        path = base_dir / path
    
    if not path.exists():
        raise FileNotFoundError(f"Access file not found: {path}")
    
    if not path.suffix.lower() == ".accdb":
        logger.warning("File does not have .accdb extension: %s", path)
    
    return path


# =============================================================================
# Database Connection
# =============================================================================

@contextmanager
def open_access_connection(file_path: Path):
    """
    Open a connection to an Access database.
    
    This is a context manager that ensures the connection is properly closed.
    
    Args:
        file_path: Path to the .accdb file.
        
    Yields:
        pyodbc.Connection object.
        
    Raises:
        ImportError: If pyodbc is not installed.
        pyodbc.Error: If connection fails.
        
    Example:
        >>> with open_access_connection(Path("data.accdb")) as conn:
        ...     cursor = conn.cursor()
        ...     cursor.execute("SELECT * FROM tblKartei")
    """
    try:
        import pyodbc
    except ImportError as e:
        raise ImportError(
            "pyodbc is required for Access database access. "
            "Install it with: pip install pyodbc"
        ) from e
    
    conn_string = build_connection_string(file_path)
    logger.info("Connecting to Access database: %s", file_path)
    
    conn = pyodbc.connect(conn_string)
    try:
        yield conn
    finally:
        conn.close()
        logger.debug("Closed connection to: %s", file_path)


# =============================================================================
# Type Definitions
# =============================================================================

# RowDict: dictionary representation of an Access row
# Keys: "ID", "Value1".."Value52", "InteriorColor1".."InteriorColor51",
#       "FontColor3", "FontColor18"
RowDict = dict[str, str | int | float | None]


# =============================================================================
# Table Reading Functions
# =============================================================================

def _build_select_query(table_name: str) -> str:
    """
    Build a SELECT query for a Kartei table.
    
    Selects all fields in the standard Kartei table structure:
    - ID
    - Value1 through Value52
    - InteriorColor1 through InteriorColor51
    - FontColor3, FontColor18
    
    Args:
        table_name: Name of the table (tblKartei, pre_tblKartei, decl_tblKartei).
        
    Returns:
        SQL SELECT statement.
    """
    # Build field list
    fields = ["ID"]
    
    # Value1..Value52
    for i in range(1, 53):
        fields.append(f"Value{i}")
    
    # InteriorColor1..InteriorColor51
    for i in range(1, 52):
        fields.append(f"InteriorColor{i}")
    
    # FontColor3, FontColor18
    fields.extend(["FontColor3", "FontColor18"])
    
    field_list = ", ".join(fields)
    return f"SELECT {field_list} FROM [{table_name}]"


def _row_to_dict(cursor_description, row) -> RowDict:
    """
    Convert a pyodbc row to a dictionary.
    
    Args:
        cursor_description: cursor.description from pyodbc
        row: A row tuple from cursor.fetchone() or fetchmany()
        
    Returns:
        RowDict with column names as keys.
    """
    return {
        col[0]: value
        for col, value in zip(cursor_description, row)
    }


def _read_table(conn, table_name: str, batch_size: int = 1000) -> Iterator[RowDict]:
    """
    Read all rows from a table as dictionaries.
    
    Args:
        conn: pyodbc connection object.
        table_name: Name of the table to read.
        batch_size: Number of rows to fetch at a time.
        
    Yields:
        RowDict for each row in the table.
        
    Raises:
        Exception: If table does not exist or query fails.
    """
    cursor = conn.cursor()
    query = _build_select_query(table_name)
    
    logger.info("Executing query for table: %s", table_name)
    logger.debug("Query: %s", query)
    
    try:
        cursor.execute(query)
    except Exception as e:
        logger.warning("Failed to query table %s: %s", table_name, e)
        raise
    
    description = cursor.description
    
    while True:
        rows = cursor.fetchmany(batch_size)
        if not rows:
            break
        
        for row in rows:
            yield _row_to_dict(description, row)
    
    cursor.close()


def _table_exists(conn, table_name: str) -> bool:
    """
    Check if a table exists in the database.
    
    Args:
        conn: pyodbc connection object.
        table_name: Name of the table to check.
        
    Returns:
        True if the table exists, False otherwise.
    """
    cursor = conn.cursor()
    try:
        # Try a simple query to check if table exists
        cursor.execute(f"SELECT TOP 1 ID FROM [{table_name}]")
        return True
    except Exception:
        return False
    finally:
        cursor.close()


# =============================================================================
# Public API: Load Functions
# =============================================================================

def load_tbl_kartei(conn, year: int) -> Iterator[RowDict]:
    """
    Load all rows from tblKartei (main Kartei table).
    
    This reads the primary source of truth table containing all approved
    records for a given year.
    
    Args:
        conn: pyodbc connection to the Access database.
        year: The year this data belongs to (for logging/context).
        
    Yields:
        RowDict for each row in tblKartei.
        
    Note:
        The year parameter is informational only; the connection should
        already be opened to the correct year's .accdb file.
    """
    logger.info("Loading tblKartei for year %d", year)
    
    if not _table_exists(conn, "tblKartei"):
        logger.warning("tblKartei does not exist in database for year %d", year)
        return
    
    count = 0
    for row in _read_table(conn, "tblKartei"):
        count += 1
        yield row
    
    logger.info("Loaded %d rows from tblKartei (year %d)", count, year)


def load_pre_tbl_kartei(conn, year: int) -> Iterator[RowDict]:
    """
    Load all rows from pre_tblKartei (pending changes table).
    
    This reads records that are awaiting Superadmin approval.
    
    Args:
        conn: pyodbc connection to the Access database.
        year: The year this data belongs to (for logging/context).
        
    Yields:
        RowDict for each row in pre_tblKartei.
        
    Note:
        The table may not exist if there are no pending changes.
    """
    logger.info("Loading pre_tblKartei for year %d", year)
    
    if not _table_exists(conn, "pre_tblKartei"):
        logger.info("pre_tblKartei does not exist in database for year %d (no pending changes)", year)
        return
    
    count = 0
    for row in _read_table(conn, "pre_tblKartei"):
        count += 1
        yield row
    
    logger.info("Loaded %d rows from pre_tblKartei (year %d)", count, year)


def load_decl_tbl_kartei(conn, year: int) -> Iterator[RowDict]:
    """
    Load all rows from decl_tblKartei (declined changes table).
    
    This reads records that have been declined by Superadmin.
    
    Args:
        conn: pyodbc connection to the Access database.
        year: The year this data belongs to (for logging/context).
        
    Yields:
        RowDict for each row in decl_tblKartei.
        
    Note:
        The table may not exist if there are no declined changes.
    """
    logger.info("Loading decl_tblKartei for year %d", year)
    
    if not _table_exists(conn, "decl_tblKartei"):
        logger.info("decl_tblKartei does not exist in database for year %d (no declined changes)", year)
        return
    
    count = 0
    for row in _read_table(conn, "decl_tblKartei"):
        count += 1
        yield row
    
    logger.info("Loaded %d rows from decl_tblKartei (year %d)", count, year)


def get_table_row_count(conn, table_name: str) -> int:
    """
    Get the number of rows in a table.
    
    Args:
        conn: pyodbc connection object.
        table_name: Name of the table.
        
    Returns:
        Number of rows, or 0 if table doesn't exist.
    """
    if not _table_exists(conn, table_name):
        return 0
    
    cursor = conn.cursor()
    cursor.execute(f"SELECT COUNT(*) FROM [{table_name}]")
    result = cursor.fetchone()
    cursor.close()
    
    return result[0] if result else 0


def get_database_stats(conn, year: int) -> dict[str, int]:
    """
    Get statistics about the tables in the database.
    
    Args:
        conn: pyodbc connection object.
        year: Year for logging.
        
    Returns:
        Dictionary with table names as keys and row counts as values.
    """
    stats = {
        "tblKartei": get_table_row_count(conn, "tblKartei"),
        "pre_tblKartei": get_table_row_count(conn, "pre_tblKartei"),
        "decl_tblKartei": get_table_row_count(conn, "decl_tblKartei"),
    }
    
    logger.info(
        "Database stats for year %d: tblKartei=%d, pre_tblKartei=%d, decl_tblKartei=%d",
        year, stats["tblKartei"], stats["pre_tblKartei"], stats["decl_tblKartei"]
    )
    
    return stats
