"""
Global display preferences for the Smart Scheduler suite.
Manages the toggle between showing postcodes/locations vs client names.
"""

import json
import os
from pathlib import Path
import tkinter as tk

# Global state
_display_preference = None
_preference_file = None
_preference_callbacks = []


def initialize(config_dir=None):
    """Initialize display preferences system
    
    Args:
        config_dir: Directory to store preferences. If None, uses current working directory.
    """
    global _preference_file, _display_preference
    
    if config_dir is None:
        config_dir = os.getcwd()
    
    _preference_file = os.path.join(config_dir, "display_preferences.json")
    print(f"[DEBUG display_prefs] Initializing with preference file: {_preference_file}")
    
    # Load existing preferences or create default
    if os.path.exists(_preference_file):
        try:
            with open(_preference_file, 'r') as f:
                data = json.load(f)
                _display_preference = data.get('show_names', False)
            print(f"[DEBUG display_prefs] Loaded preferences: show_names = {_display_preference}")
        except Exception as e:
            print(f"[DEBUG display_prefs] Error loading preferences: {e}")
            _display_preference = False
    else:
        _display_preference = False
        print(f"[DEBUG display_prefs] No existing preference file, using default: False")
    
    return _display_preference


def get_show_names():
    """Get current preference: True = show names, False = show postcodes"""
    global _display_preference
    if _display_preference is None:
        initialize()
    print(f"[DEBUG display_prefs] get_show_names() returning: {_display_preference}")
    return _display_preference


def set_show_names(show_names):
    """Set the display preference and persist to file"""
    global _display_preference
    print(f"[DEBUG display_prefs] set_show_names() called with: {show_names}")
    _display_preference = show_names
    
    if _preference_file:
        try:
            with open(_preference_file, 'w') as f:
                json.dump({'show_names': show_names}, f)
            print(f"[DEBUG display_prefs] Preferences saved to {_preference_file}")
        except Exception as e:
            print(f"[DEBUG display_prefs] Warning: Could not save display preference: {e}")
    
    # Notify all listeners
    print(f"[DEBUG display_prefs] Notifying {len(_preference_callbacks)} callbacks")
    for callback in _preference_callbacks:
        try:
            callback(show_names)
            print(f"[DEBUG display_prefs] Callback executed successfully")
        except Exception as e:
            print(f"[DEBUG display_prefs] Warning: Callback error: {e}")


def register_callback(callback):
    """Register a callback to be called when preference changes
    
    Args:
        callback: Function that takes one argument (show_names: bool)
    """
    _preference_callbacks.append(callback)


def unregister_callback(callback):
    """Unregister a callback"""
    if callback in _preference_callbacks:
        _preference_callbacks.remove(callback)


def format_location(postcode, client_name=None, show_names=None):
    """Format a location for display based on current preference
    
    Args:
        postcode: The postcode/location identifier
        client_name: Optional client name. If None, only postcode is used.
        show_names: Override the global setting (for testing). If None, uses global.
    
    Returns:
        Formatted string for display
    """
    if show_names is None:
        show_names = get_show_names()
    
    # If showing names and client name exists, use it
    if show_names and client_name:
        return str(client_name)
    
    # Otherwise return postcode in bold/italics
    return f"***{postcode}***"  # Will be formatted differently in UI (bold/italic)


def format_location_raw(postcode, client_name=None, show_names=None):
    """Get raw text (without formatting markers) for use in data structures
    
    Args:
        postcode: The postcode/location identifier
        client_name: Optional client name
        show_names: Override the global setting
    
    Returns:
        Plain text to display
    """
    if show_names is None:
        show_names = get_show_names()
    
    if show_names and client_name:
        return str(client_name)
    
    return str(postcode)


def get_location_from_data(row, show_names=None):
    """Extract and format location from a data row
    
    Args:
        row: Dict or pandas Series with 'postcode' and optional 'client_name' keys
        show_names: Override the global setting
    
    Returns:
        Formatted location string
    """
    postcode = str(row.get('postcode', '')) if hasattr(row, 'get') else str(row['postcode'])
    client_name = row.get('client_name', None) if hasattr(row, 'get') else (row['client_name'] if 'client_name' in row else None)
    
    if client_name and pd.notna(client_name):
        client_name = str(client_name).strip()
        if not client_name:
            client_name = None
    else:
        client_name = None
    
    return format_location(postcode, client_name, show_names)


# Check for pandas
try:
    import pandas as pd
except ImportError:
    pd = None
