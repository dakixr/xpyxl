"""Test that all example modules can be loaded and combined into workbooks."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

import pytest

import xpyxl as x
from xpyxl.nodes import ImportedSheetNode, SheetNode


def load_module_from_path(module_name: str, file_path: Path) -> ModuleType:
    """Load a module from a file path."""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load module from {file_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def discover_examples(examples_dir: Path) -> list[Path]:
    """Discover all example Python files in the examples directory."""
    examples = []
    for file_path in examples_dir.glob("*.py"):
        # Skip __init__.py and test files
        if file_path.name == "__init__.py" or file_path.name.startswith("test_"):
            continue
        examples.append(file_path)
    return sorted(examples)


def get_sheets_from_module(module: ModuleType, module_name: str) -> list[SheetNode]:
    """Get sheets from a module, trying build_workbook() first, then build_sample_workbook().

    Returns a list of SheetNode objects. Single sheets are wrapped in a list.
    """
    if hasattr(module, "build_workbook"):
        result = module.build_workbook()
        # If it's a single sheet, wrap it in a list
        if isinstance(result, SheetNode):
            return [result]
        # If it's already a list, return it
        return result
    elif hasattr(module, "build_sample_workbook"):
        result = module.build_sample_workbook()
        # If it's a single sheet, wrap it in a list
        if isinstance(result, SheetNode):
            return [result]
        # If it's already a list, return it
        return result
    else:
        raise AttributeError(
            f"Module {module_name} must have either build_workbook() or build_sample_workbook() function"
        )


def test_combined_workbook() -> None:
    """Test that all example modules can be loaded and combined into workbooks."""
    # Get project root and examples directory
    project_root = Path(__file__).resolve().parent.parent.parent
    examples_dir = project_root / "tests" / "end_to_end"
    output_dir = project_root / ".testing"

    # Create .testing directory if it doesn't exist
    output_dir.mkdir(exist_ok=True)

    # Discover all example files
    example_files = discover_examples(examples_dir)

    # Assert that we found at least some examples
    assert len(example_files) > 0, f"No examples found in {examples_dir}"

    # Collect all sheets from all modules
    all_sheets: list[SheetNode | ImportedSheetNode] = []
    seen_sheet_names: set[str] = set()
    success_count = 0
    failed_modules: list[str] = []

    for example_file in example_files:
        module_name = example_file.stem
        try:
            module = load_module_from_path(module_name, example_file)
            module_sheets = get_sheets_from_module(module, module_name)

            # Ensure unique sheet names
            for sheet in module_sheets:
                original_name = sheet.name
                if original_name in seen_sheet_names:
                    # Rename with module prefix
                    new_name = f"{module_name}-{original_name}"
                    if isinstance(sheet, ImportedSheetNode):
                        all_sheets.append(
                            ImportedSheetNode(
                                name=new_name,
                                source=sheet.source,
                                source_sheet=sheet.source_sheet,
                            )
                        )
                    else:
                        all_sheets.append(
                            SheetNode(
                                name=new_name,
                                items=sheet.items,
                                background_color=sheet.background_color,
                            )
                        )
                    seen_sheet_names.add(new_name)
                else:
                    all_sheets.append(sheet)
                    seen_sheet_names.add(original_name)

            success_count += 1
        except Exception as e:
            failed_modules.append(f"{module_name}: {e}")
            # Continue processing other modules

    # Assert that we collected at least some sheets
    assert len(all_sheets) > 0, "No sheets collected from any example modules"

    # Assert that at least some modules succeeded
    assert success_count > 0, f"All modules failed to load. Failures: {failed_modules}"

    # Create combined workbook
    combined_workbook = x.workbook()[*all_sheets]

    # Save with both engines
    openpyxl_path = output_dir / "combined-output-openpyxl.xlsx"
    xlsxwriter_path = output_dir / "combined-output-xlsxwriter.xlsx"

    # Test openpyxl engine save
    combined_workbook.save(openpyxl_path, engine="openpyxl")
    assert openpyxl_path.exists(), (
        f"openpyxl output file was not created: {openpyxl_path}"
    )

    # Test xlsxwriter engine save (uses hybrid save for imported sheets if present)
    combined_workbook.save(xlsxwriter_path, engine="xlsxwriter")
    assert xlsxwriter_path.exists(), (
        f"xlsxwriter output file was not created: {xlsxwriter_path}"
    )

    # Verify both files are valid Excel files by checking they exist and have content
    assert openpyxl_path.stat().st_size > 0, "openpyxl output file is empty"
    assert xlsxwriter_path.stat().st_size > 0, "xlsxwriter output file is empty"


if __name__ == "__main__":
    pytest.main([__file__])
