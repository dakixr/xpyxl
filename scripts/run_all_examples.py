"""Run all example modules and save outputs to .testing folder."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType


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


def derive_output_filename(module_name: str) -> str:
    """Derive output filename from module name."""
    # Convert underscores to hyphens and append -output.xlsx
    return f"{module_name.replace('_', '-')}-output.xlsx"


def main() -> None:
    """Run all example modules and save outputs to .testing folder."""
    # Get project root and examples directory
    project_root = Path(__file__).resolve().parent.parent
    examples_dir = project_root / "examples"
    output_dir = project_root / ".testing"
    src_dir = project_root / "src"

    # Add src directory to Python path so xpyxl can be imported
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))

    # Create .testing directory if it doesn't exist
    output_dir.mkdir(exist_ok=True)

    # Discover all example files
    example_files = discover_examples(examples_dir)

    if not example_files:
        print(f"No examples found in {examples_dir}")
        return

    print(f"Found {len(example_files)} example(s) in {examples_dir}")
    print(f"Saving outputs to {output_dir.resolve()}")
    print("-" * 60)

    success_count = 0
    for example_file in example_files:
        module_name = example_file.stem
        output_filename = derive_output_filename(module_name)
        output_path = output_dir / output_filename

        print(f"\nRunning {module_name}...")
        try:
            module = load_module_from_path(module_name, example_file)
            module.main(output_path=output_path)
            print(f"✓ Successfully generated {output_filename}")
            success_count += 1
        except Exception as e:
            print(f"✗ Error running {module_name}: {e}")
            import traceback

            traceback.print_exc()

    print("\n" + "-" * 60)
    print(f"Completed: {success_count}/{len(example_files)} examples succeeded")
    print(f"Outputs saved to {output_dir.resolve()}")


if __name__ == "__main__":
    main()
