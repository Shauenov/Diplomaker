# -*- coding: utf-8 -*-
"""
Custom Exceptions
=================
Exception hierarchy for diploma automation system.
"""


class DiplomaAutomationError(Exception):
    """
    Base exception for all diploma automation errors.
    
    Use this for catching any error from the system.
    """
    pass


class ConfigurationError(DiplomaAutomationError):
    """
    Raised when configuration is invalid or missing.
    
    Examples:
        - Invalid program code
        - Missing source file path
        - Invalid language code
    """
    pass


class ParseError(DiplomaAutomationError):
    """
    Raised when parsing source data fails.
    
    Examples:
        - Cannot read Excel file
        - Excel structure doesn't match expected format
        - Invalid cell value format
    """
    pass


class ValidationError(DiplomaAutomationError):
    """
    Raised when data validation fails.
    
    Examples:
        - Missing required student data
        - Invalid grade value (out of 0-100 range)
        - Incomplete subject requirements
    """
    pass


class GenerationError(DiplomaAutomationError):
    """
    Raised when diploma generation fails.
    
    Examples:
        - Cannot create output file
        - Cannot write to Excel
        - Template formatting error
    """
    pass
