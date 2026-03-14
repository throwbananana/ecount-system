import pandas as pd

from local_llm_analyzer import LocalLLMAnalyzer


def test_related_context_for_expense_and_sales_sheet():
    analyzer = LocalLLMAnalyzer()
    analyzer._prepare_related_cache = lambda _: {
        "kpi_pack": "KPI",
        "expense_pack": "EXPENSE",
        "sales_pack": "SALES",
    }

    expense_ctx = analyzer._build_related_context_for_sheet("dummy.xlsx", "费用对比", pd.DataFrame())
    assert "KPI" in expense_ctx
    assert "EXPENSE" in expense_ctx
    assert "SALES" not in expense_ctx

    sales_ctx = analyzer._build_related_context_for_sheet("dummy.xlsx", "产品贡献毛利", pd.DataFrame())
    assert "KPI" in sales_ctx
    assert "SALES" in sales_ctx


if __name__ == "__main__":
    test_related_context_for_expense_and_sales_sheet()
    print("PASS: test_local_llm_related_context")
