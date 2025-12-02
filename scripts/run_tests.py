"""Run all example modules and save outputs to .testing folder."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

# Add src directory to Python path before importing xpyxl
_project_root = Path(__file__).resolve().parent.parent
_src_dir = _project_root / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

import xpyxl as x  # noqa: E402
from xpyxl.nodes import SheetNode  # noqa: E402


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


def main() -> None:
    """Run all example modules and save combined outputs to .testing folder."""
    # Get project root and examples directory
    project_root = Path(__file__).resolve().parent.parent
    examples_dir = project_root / "tests"
    output_dir = project_root / ".testing"

    # Create .testing directory if it doesn't exist
    output_dir.mkdir(exist_ok=True)

    # Discover all example files
    example_files = discover_examples(examples_dir)

    if not example_files:
        print(f"No examples found in {examples_dir}")
        return

    print(f"Found {len(example_files)} example(s) in {examples_dir}")
    print(f"Saving combined output to {output_dir.resolve()}")
    print("-" * 60)

    # Collect all sheets from all modules
    all_sheets: list[SheetNode] = []
    seen_sheet_names: set[str] = set()
    success_count = 0

    for example_file in example_files:
        module_name = example_file.stem
        print(f"\nLoading {module_name}...")
        try:
            module = load_module_from_path(module_name, example_file)
            module_sheets = get_sheets_from_module(module, module_name)

            # Ensure unique sheet names
            for sheet in module_sheets:
                original_name = sheet.name
                if original_name in seen_sheet_names:
                    # Rename with module prefix
                    new_name = f"{module_name}-{original_name}"
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

            print(
                f"✓ Successfully loaded {len(module_sheets)} sheet(s) from {module_name}"
            )
            success_count += 1
        except Exception as e:
            print(f"✗ Error loading {module_name}: {e}")
            import traceback

            traceback.print_exc()

    if not all_sheets:
        print("\nNo sheets collected. Cannot create combined workbook.")
        return

    # Create combined workbook
    print(f"\nCreating combined workbook with {len(all_sheets)} sheet(s)...")
    combined_workbook = x.workbook()[*all_sheets]

    # Save with both engines
    openpyxl_path = output_dir / "combined-output-openpyxl.xlsx"
    xlsxwriter_path = output_dir / "combined-output-xlsxwriter.xlsx"

    print(f"\nSaving with openpyxl engine to {openpyxl_path.name}...")
    try:
        combined_workbook.save(openpyxl_path, engine="openpyxl")
        print(f"✓ Successfully saved {openpyxl_path.name}")
    except Exception as e:
        print(f"✗ Error saving with openpyxl: {e}")
        import traceback

        traceback.print_exc()

    print(f"\nSaving with xlsxwriter engine to {xlsxwriter_path.name}...")
    try:
        combined_workbook.save(xlsxwriter_path, engine="xlsxwriter")
        print(f"✓ Successfully saved {xlsxwriter_path.name}")
    except Exception as e:
        print(f"✗ Error saving with xlsxwriter: {e}")
        import traceback

        traceback.print_exc()

    print("\n" + "-" * 60)
    print(f"Completed: {success_count}/{len(example_files)} examples succeeded")
    print(f"Combined outputs saved to {output_dir.resolve()}")


if __name__ == "__main__":
    main()
