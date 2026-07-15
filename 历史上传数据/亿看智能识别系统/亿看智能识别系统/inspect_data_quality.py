
import os
import pandas as pd
from report_generator import ReportGenerator

def inspect_quality():
    base_dir = "基础资料"
    if not os.path.exists(base_dir):
        # Try full path if relative fails
        base_dir = r"C:\Users\123\Downloads\亿看智能识别系统\基础资料"
        
    print(f"Target directory: {base_dir}")
    generator = ReportGenerator(base_dir)
    generator.load_all_data()
    
    issues = generator.data_quality_issues
    print(f"\nTotal issues found: {len(issues)}")
    
    # Group by severity
    by_severity = {"ERROR": [], "WARN": [], "INFO": []}
    for issue in issues:
        severity = issue.get("severity", "WARN").upper()
        by_severity.setdefault(severity, []).append(issue)
        
    print("\n--- ERRORs ---")
    for issue in by_severity["ERROR"]:
        print(f"[{issue.get('category')}] [{issue.get('period')}] {issue.get('issue_type')}: {issue.get('detail')}")
        
    print("\n--- TOP 10 WARNs ---")
    for issue in by_severity["WARN"][:10]:
         print(f"[{issue.get('category')}] [{issue.get('period')}] {issue.get('issue_type')}: {issue.get('detail')}")

    if len(by_severity["WARN"]) > 10:
        print(f"... and {len(by_severity['WARN']) - 10} more warnings.")

    print("\n--- TOP 5 INFOs ---")
    for issue in by_severity["INFO"][:5]:
         print(f"[{issue.get('category')}] [{issue.get('period')}] {issue.get('issue_type')}: {issue.get('detail')}")

if __name__ == "__main__":
    inspect_quality()
