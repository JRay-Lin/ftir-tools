import sys
from pathlib import Path
import tomllib


def get_version():
    """
    Get application version from pyproject.toml.
    Fallback to a default version if the file cannot be read.
    """
    try:
        # Get the project root directory
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller executable
            app_dir = Path(sys.executable).parent
        else:
            # Running as Python script
            app_dir = Path(__file__).parent.parent
        
        pyproject_path = app_dir / "pyproject.toml"
        
        if pyproject_path.exists():
            with open(pyproject_path, "rb") as f:
                data = tomllib.load(f)
                return data.get("project", {}).get("version", "Unknown")
        else:
            return "Unknown"
    except Exception:
        return "Unknown"


def get_app_info():
    """
    Get application information from pyproject.toml.
    """
    try:
        # Get the project root directory
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller executable
            app_dir = Path(sys.executable).parent
        else:
            # Running as Python script
            app_dir = Path(__file__).parent.parent
        
        pyproject_path = app_dir / "pyproject.toml"
        
        if pyproject_path.exists():
            with open(pyproject_path, "rb") as f:
                data = tomllib.load(f)
                project_info = data.get("project", {})
                return {
                    "name": project_info.get("name", "FTIR Tools"),
                    "version": project_info.get("version", "Unknown"),
                    "description": project_info.get("description", "FTIR Data Analysis Tool"),
                    "python_requirement": project_info.get("requires-python", ">=3.12")
                }
        else:
            return {
                "name": "FTIR Tools",
                "version": "Unknown", 
                "description": "FTIR Data Analysis Tool",
                "python_requirement": ">=3.12"
            }
    except Exception:
        return {
            "name": "FTIR Tools",
            "version": "Unknown",
            "description": "FTIR Data Analysis Tool", 
            "python_requirement": ">=3.12"
        }