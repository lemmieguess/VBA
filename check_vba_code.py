#!/usr/bin/env python3
"""
VBA Code Validator - Checks for Common BWS Coding Errors
Usage: python check_vba_code.py <filename.bas>
Example: python check_vba_code.py BWS_v1.7.10.bas
"""

import sys
import re
from pathlib import Path


class Colors:
    """ANSI color codes for terminal output"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    BOLD = '\033[1m'
    END = '\033[0m'


def print_header(text):
    """Print a formatted header"""
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'=' * 60}{Colors.END}")
    print(f"{Colors.BOLD}{Colors.BLUE}{text:^60}{Colors.END}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'=' * 60}{Colors.END}\n")


def print_success(text):
    """Print success message"""
    print(f"{Colors.GREEN}✅ {text}{Colors.END}")


def print_error(text):
    """Print error message"""
    print(f"{Colors.RED}❌ {text}{Colors.END}")


def print_warning(text):
    """Print warning message"""
    print(f"{Colors.YELLOW}⚠️  {text}{Colors.END}")


def check_line_continuations(lines):
    """Check for excessive line continuations (max 20 recommended)"""
    print_header("Checking Line Continuations")

    issues = []
    continuation_count = 0
    start_line = 0

    for i, line in enumerate(lines, 1):
        if line.rstrip().endswith(' _'):
            if continuation_count == 0:
                start_line = i
            continuation_count += 1
        else:
            if continuation_count > 20:
                issues.append({
                    'start': start_line,
                    'end': i,
                    'count': continuation_count
                })
            continuation_count = 0

    if issues:
        print_error(f"Found {len(issues)} section(s) with too many line continuations:")
        for issue in issues:
            print(f"   Lines {issue['start']}-{issue['end']}: {issue['count']} continuations")
            print(f"   {Colors.YELLOW}Recommended: Split into multiple variables (max 20){Colors.END}")
        return False
    else:
        print_success("All line continuations are within safe limits (<20)")
        return True


def check_dim_placement(lines):
    """Check for Dim statements inside loops or If blocks"""
    print_header("Checking Variable Declaration Placement")

    issues = []
    in_function = False
    function_name = ""
    indent_stack = []
    function_start = 0

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        # Track function boundaries
        if stripped.startswith(('Sub ', 'Function ', 'Public Sub', 'Private Sub',
                               'Public Function', 'Private Function')):
            in_function = True
            function_name = stripped.split('(')[0].replace('Sub ', '').replace('Function ', '').strip()
            function_start = i
            indent_stack = []
            continue

        if stripped.startswith(('End Sub', 'End Function')):
            in_function = False
            continue

        if not in_function:
            continue

        # Track block depth
        if any(stripped.startswith(x) for x in ['For ', 'If ', 'While ', 'Do ', 'With ',
                                                  'Select Case', 'ElseIf ']):
            indent_stack.append((i, stripped.split()[0]))

        if any(stripped.startswith(x) for x in ['Next', 'End If', 'Wend', 'Loop',
                                                  'End With', 'End Select']):
            if indent_stack:
                indent_stack.pop()

        # Check for Dim inside blocks (but allow at function level)
        if stripped.startswith('Dim '):
            # Calculate if we're inside a block (not just at function level)
            lines_from_function_start = i - function_start

            if indent_stack:  # We're inside a block
                issues.append({
                    'line': i,
                    'text': line.strip(),
                    'function': function_name,
                    'block': indent_stack[-1][1] if indent_stack else 'Unknown',
                    'block_line': indent_stack[-1][0] if indent_stack else 'Unknown'
                })

    if issues:
        print_error(f"Found {len(issues)} Dim statement(s) in wrong location:")
        for issue in issues:
            print(f"\n   {Colors.RED}Line {issue['line']} in {issue['function']}():{Colors.END}")
            print(f"      {issue['text']}")
            print(f"      Inside: {issue['block']} block (started at line {issue['block_line']})")
            print(f"      {Colors.YELLOW}FIX: Move this Dim to the top of {issue['function']}(){Colors.END}")
        return False
    else:
        print_success("All Dim statements are properly placed at function level")
        return True


def check_version_consistency(lines):
    """Check that version numbers are consistent"""
    print_header("Checking Version Number Consistency")

    version_pattern = r'v?\d+\.\d+\.\d+'
    versions_found = {}

    for i, line in enumerate(lines, 1):
        # Check BWS_VERSION constant (must be a Const declaration, not just any use)
        if 'Const BWS_VERSION' in line and '=' in line:
            match = re.search(r'"([^"]+)"', line)
            if match:
                versions_found['BWS_VERSION'] = (match.group(1), i)

        # Check BWS_About header
        if '================' in line and 'BWS v' in line:
            match = re.search(version_pattern, line)
            if match:
                versions_found['About_Header'] = (match.group(0), i)

    if len(versions_found) < 2:
        print_warning("Could not find all version number locations")
        return True

    # Check consistency
    versions = [v[0] for v in versions_found.values()]
    if len(set(versions)) == 1:
        print_success(f"All version numbers consistent: {versions[0]}")
        for location, (version, line_no) in versions_found.items():
            print(f"   Line {line_no}: {location} = {version}")
        return True
    else:
        print_error("Version numbers are INCONSISTENT:")
        for location, (version, line_no) in versions_found.items():
            print(f"   Line {line_no}: {location} = {version}")
        print(f"\n   {Colors.YELLOW}FIX: Update all version numbers to match{Colors.END}")
        return False


def check_common_errors(lines):
    """Check for other common VBA errors"""
    print_header("Checking for Common VBA Errors")

    issues = []

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        # Check for .Index on Paragraph objects (doesn't exist!)
        if '.Index' in line and 'Paragraph' in line:
            issues.append({
                'line': i,
                'type': 'Paragraph.Index',
                'text': line.strip(),
                'fix': 'Use Range.MoveEnd() instead - Paragraph objects have no .Index property'
            })

        # Check for very long lines (>1024 chars)
        if len(line) > 1000:
            issues.append({
                'line': i,
                'type': 'Long Line',
                'text': line[:100] + '...',
                'fix': 'Break into multiple lines using " _" continuation'
            })

    if issues:
        print_warning(f"Found {len(issues)} potential issue(s):")
        for issue in issues:
            print(f"\n   Line {issue['line']}: {issue['type']}")
            print(f"      {issue['text']}")
            print(f"      {Colors.YELLOW}FIX: {issue['fix']}{Colors.END}")
        return True  # These are warnings, not critical
    else:
        print_success("No common VBA errors detected")
        return True


def main():
    """Main validation function"""
    if len(sys.argv) < 2:
        print(f"{Colors.RED}Usage: python check_vba_code.py <filename.bas>{Colors.END}")
        print(f"{Colors.YELLOW}Example: python check_vba_code.py BWS_v1.7.10.bas{Colors.END}")
        sys.exit(1)

    filename = sys.argv[1]
    filepath = Path(filename)

    if not filepath.exists():
        print(f"{Colors.RED}Error: File '{filename}' not found{Colors.END}")
        sys.exit(1)

    print(f"\n{Colors.BOLD}Validating: {filename}{Colors.END}")

    # Read file
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"{Colors.RED}Error reading file: {e}{Colors.END}")
        sys.exit(1)

    # Run all checks
    results = []
    results.append(check_line_continuations(lines))
    results.append(check_dim_placement(lines))
    results.append(check_version_consistency(lines))
    results.append(check_common_errors(lines))

    # Summary
    print_header("Validation Summary")

    passed = sum(results)
    total = len(results)

    if all(results):
        print(f"{Colors.GREEN}{Colors.BOLD}🎉 ALL CHECKS PASSED! ({passed}/{total}){Colors.END}")
        print(f"{Colors.GREEN}Your code is ready to commit!{Colors.END}")
        sys.exit(0)
    else:
        print(f"{Colors.RED}{Colors.BOLD}⚠️  SOME CHECKS FAILED ({passed}/{total} passed){Colors.END}")
        print(f"{Colors.YELLOW}Please fix the issues above before committing{Colors.END}")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{Colors.YELLOW}Validation cancelled{Colors.END}")
        sys.exit(1)
    except Exception as e:
        print(f"\n{Colors.RED}Unexpected error: {e}{Colors.END}")
        sys.exit(1)
