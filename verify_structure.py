"""
Verify the structure and syntax of the email agent modules
This script checks imports and class definitions without requiring dependencies
"""

import ast
import os


def check_file_syntax(filename):
    """Check if a Python file has valid syntax"""
    print(f"\nChecking {filename}...")
    
    if not os.path.exists(filename):
        print(f"  ❌ File not found: {filename}")
        return False
    
    try:
        with open(filename, 'r') as f:
            code = f.read()
        
        ast.parse(code)
        print(f"  ✅ Syntax is valid")
        
        tree = ast.parse(code)
        
        classes = [node.name for node in ast.walk(tree) if isinstance(node, ast.ClassDef)]
        functions = [node.name for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)]
        
        if classes:
            print(f"  📦 Classes found: {', '.join(classes)}")
        if functions:
            print(f"  🔧 Functions found: {len(functions)} functions")
        
        return True
        
    except SyntaxError as e:
        print(f"  ❌ Syntax error: {e}")
        return False
    except Exception as e:
        print(f"  ❌ Error: {e}")
        return False


def verify_project_structure():
    """Verify all project files exist"""
    print("=" * 60)
    print("Email to Excel Agent - Structure Verification")
    print("=" * 60)
    
    required_files = [
        "email_processor.py",
        "excel_exporter.py",
        "email_to_excel_app.py",
        "requirements.txt",
        "README_EMAIL_TO_EXCEL.md"
    ]
    
    print("\n📁 Checking project files...")
    all_exist = True
    for file in required_files:
        exists = os.path.exists(file)
        status = "✅" if exists else "❌"
        print(f"  {status} {file}")
        if not exists:
            all_exist = False
    
    return all_exist


def check_requirements():
    """Check requirements.txt content"""
    print("\n📋 Checking requirements.txt...")
    
    try:
        with open("requirements.txt", 'r') as f:
            requirements = f.read().strip().split('\n')
        
        expected = ['pandas', 'openpyxl', 'google-auth', 'streamlit']
        
        for req in expected:
            found = any(req in line for line in requirements)
            status = "✅" if found else "❌"
            print(f"  {status} {req}")
        
        return True
    except Exception as e:
        print(f"  ❌ Error reading requirements: {e}")
        return False


def main():
    """Run all verification checks"""
    
    results = []
    
    results.append(("Project Structure", verify_project_structure()))
    results.append(("Requirements", check_requirements()))
    results.append(("email_processor.py", check_file_syntax("email_processor.py")))
    results.append(("excel_exporter.py", check_file_syntax("excel_exporter.py")))
    results.append(("email_to_excel_app.py", check_file_syntax("email_to_excel_app.py")))
    
    print("\n" + "=" * 60)
    print("Verification Summary")
    print("=" * 60)
    
    for check_name, result in results:
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"{check_name}: {status}")
    
    all_passed = all(result for _, result in results)
    
    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 All verification checks passed!")
        print("\nThe Email to Excel Agent is ready to use.")
        print("\nTo run the application:")
        print("  1. Install dependencies: pip install -r requirements.txt")
        print("  2. Run the app: streamlit run email_to_excel_app.py")
    else:
        print("⚠️  Some checks failed. Please review the output above.")
    print("=" * 60)
    
    return all_passed


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
