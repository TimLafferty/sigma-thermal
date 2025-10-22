"""
Parse VBA function signatures from extracted VBA code.

This script analyzes the extracted VBA files and creates a structured
inventory of all functions with their signatures and documentation.
"""

import re
import json
from pathlib import Path
from typing import Dict, List, Tuple
from dataclasses import dataclass, asdict


@dataclass
class VBAFunction:
    """Represents a VBA function signature"""
    name: str
    module: str
    return_type: str
    parameters: List[Dict[str, str]]
    visibility: str  # Public, Private, or Friend
    is_function: bool  # True for Function, False for Sub
    source_file: str


class VBAParser:
    """Parse VBA code to extract function signatures"""

    FUNCTION_PATTERN = re.compile(
        r'^(Public |Private |Friend )?(Function|Sub)\s+(\w+)\s*\((.*?)\)',
        re.MULTILINE
    )

    PARAM_PATTERN = re.compile(
        r'(\w+)\s+As\s+(\w+)|(\w+)\s*(?:,|$)'
    )

    def __init__(self, vba_file: Path):
        self.vba_file = vba_file
        self.content = vba_file.read_text(encoding='utf-8', errors='ignore')
        self.functions: List[VBAFunction] = []
        self.current_module = ""

    def parse(self) -> List[VBAFunction]:
        """Parse all functions from the VBA file"""
        lines = self.content.split('\n')

        for i, line in enumerate(lines):
            # Track current module
            if line.startswith('VBA MACRO'):
                parts = line.split()
                if len(parts) >= 3:
                    self.current_module = parts[2]

            # Match function/sub declaration
            match = self.FUNCTION_PATTERN.match(line)
            if match:
                visibility = (match.group(1) or 'Public').strip()
                func_type = match.group(2)
                func_name = match.group(3)
                params_str = match.group(4)

                # Parse parameters
                parameters = self._parse_parameters(params_str)

                # Determine return type (from following lines if Function)
                return_type = self._extract_return_type(lines, i, func_name)

                func = VBAFunction(
                    name=func_name,
                    module=self.current_module,
                    return_type=return_type,
                    parameters=parameters,
                    visibility=visibility,
                    is_function=(func_type == 'Function'),
                    source_file=self.vba_file.name
                )

                self.functions.append(func)

        return self.functions

    def _parse_parameters(self, params_str: str) -> List[Dict[str, str]]:
        """Parse function parameters"""
        if not params_str.strip():
            return []

        parameters = []
        # Split by comma but handle optional/byval/byref
        param_parts = params_str.split(',')

        for part in param_parts:
            part = part.strip()
            if not part:
                continue

            # Remove Optional, ByVal, ByRef keywords
            part = re.sub(r'(Optional |ByVal |ByRef )', '', part)

            # Match parameter name and type
            if ' As ' in part:
                name, ptype = part.split(' As ', 1)
                parameters.append({
                    'name': name.strip(),
                    'type': ptype.strip()
                })
            else:
                # Parameter without type (variant)
                parameters.append({
                    'name': part.strip(),
                    'type': 'Variant'
                })

        return parameters

    def _extract_return_type(self, lines: List[str], func_line: int, func_name: str) -> str:
        """Extract return type from function"""
        # Look at the function line for " As Type"
        func_def = lines[func_line]
        as_match = re.search(r'\)\s+As\s+(\w+)', func_def)
        if as_match:
            return as_match.group(1)

        # Check next few lines for assignment that might indicate type
        for i in range(func_line + 1, min(func_line + 10, len(lines))):
            if f'{func_name} =' in lines[i]:
                return 'Variant'

        return 'Variant'


def generate_function_inventory(vba_files: List[Path], output_file: Path):
    """Generate JSON inventory of all VBA functions"""
    all_functions = []

    for vba_file in vba_files:
        print(f"Parsing {vba_file.name}...")
        parser = VBAParser(vba_file)
        functions = parser.parse()
        all_functions.extend(functions)
        print(f"  Found {len(functions)} functions")

    # Convert to dict for JSON
    inventory = {
        'total_functions': len(all_functions),
        'functions': [asdict(f) for f in all_functions]
    }

    # Group by module
    by_module = {}
    for func in all_functions:
        module = func.module
        if module not in by_module:
            by_module[module] = []
        by_module[module].append(func.name)

    inventory['by_module'] = by_module

    # Save to JSON
    with open(output_file, 'w') as f:
        json.dump(inventory, f, indent=2)

    print(f"\nInventory saved to {output_file}")
    print(f"Total functions: {len(all_functions)}")
    print(f"Modules: {len(by_module)}")


def generate_markdown_report(inventory_file: Path, output_file: Path):
    """Generate markdown report from inventory"""
    with open(inventory_file) as f:
        inventory = json.load(f)

    with open(output_file, 'w') as f:
        f.write("# VBA Function Inventory\n\n")
        f.write(f"**Total Functions**: {inventory['total_functions']}\n\n")

        f.write("## Functions by Module\n\n")
        for module, funcs in sorted(inventory['by_module'].items()):
            f.write(f"### {module} ({len(funcs)} functions)\n\n")
            for func_data in inventory['functions']:
                if func_data['module'] == module and func_data['is_function']:
                    params = ', '.join([f"{p['name']}: {p['type']}" for p in func_data['parameters']])
                    f.write(f"- **{func_data['name']}**({params}) → {func_data['return_type']}\n")
            f.write("\n")


if __name__ == '__main__':
    # Parse VBA files
    extracted_dir = Path(__file__).parent.parent / 'extracted'
    vba_files = [
        extracted_dir / 'Engineering-Functions.vba',
        extracted_dir / 'HC2-Calculators.vba'
    ]

    # Generate inventory
    inventory_file = extracted_dir / 'function_inventory.json'
    generate_function_inventory(vba_files, inventory_file)

    # Generate markdown report
    report_file = extracted_dir / 'VBA_FUNCTION_INVENTORY.md'
    generate_markdown_report(inventory_file, report_file)

    print(f"Report saved to {report_file}")
