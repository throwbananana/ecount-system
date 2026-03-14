import os
import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl import Workbook

from report_generator import ReportGenerator


def _series_title_token(series):
    title = getattr(series, "title", None)
    if title is None:
        return None
    if getattr(title, "v", None):
        return str(title.v)

    str_ref = getattr(title, "strRef", None)
    if str_ref is None:
        tx = getattr(title, "tx", None)
        str_ref = getattr(tx, "strRef", None) if tx is not None else None
    if str_ref is None:
        return None
    return str_ref.f


def test_fill_product_summary_aggregates_rows():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {'MonthStr': '2025-12', '品目名': 'A', '数量': 2, '合计': 100, '品目编码': '001'},
        {'MonthStr': '2025-12', '品目名': 'A', '数量': 3, '合计': 180, '品目编码': '001'},
        {'MonthStr': '2025-12', '品目名': 'B', '数量': 4, '合计': 160, '品目编码': '002'},
    ])
    gen.data['cost']['2025-12'] = pd.DataFrame({
        '品目编码': ['001', '002'],
        'dummy_减少.1': [40, 20],
    })

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按产品汇总(含合计数)'
    headers = [
        '产品',
        '2025-12_销售收入',
        '2025-12_销售数量',
        '2025-12_销售成本',
        '2025-12_销售利润',
        '2025-12_毛利率',
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    ws.cell(row=2, column=1).value = 'A'
    ws.cell(row=3, column=1).value = 'B'
    ws.cell(row=4, column=1).value = '合计'

    gen._fill_product_summary(ws, '2025', '12', 'current')

    # A: revenue=280 qty=5 cost=200 profit=80 margin=80/280
    assert ws.cell(2, 2).value == 280
    assert ws.cell(2, 3).value == 5
    assert ws.cell(2, 4).value == 200
    assert ws.cell(2, 5).value == 80
    assert abs(ws.cell(2, 6).value - (80 / 280)) < 1e-12

    # Total: A + B => revenue=440 qty=9 cost=280 profit=160
    assert ws.cell(4, 2).value == 440
    assert ws.cell(4, 3).value == 9
    assert ws.cell(4, 4).value == 280
    assert ws.cell(4, 5).value == 160


def test_list_available_months_uses_core_intersection_when_loaded():
    gen = ReportGenerator('.')
    gen.data['profit'] = {'2025-12': pd.DataFrame(), '2026-01': pd.DataFrame()}
    gen.data['cost'] = {'2025-12': pd.DataFrame()}
    gen.data['asset'] = {'2025-12': pd.DataFrame()}
    gen.data['expense'] = {'2026-01': pd.DataFrame()}
    gen.data['sales'] = {'2025-12': pd.DataFrame(), '2026-01': pd.DataFrame()}

    assert gen.list_available_months() == ['2025-12']
    assert gen.list_available_years() == [2025]


def test_check_data_completeness_includes_sales_and_ar():
    gen = ReportGenerator('.')
    key = '2025-12'
    for cat in ['profit', 'cost', 'expense', 'asset', 'sales']:
        gen.data[cat][key] = pd.DataFrame({'x': [1]})

    missing = gen.check_data_completeness('2025', '12')
    assert missing == ['ar']

    gen.ar_detail_df = pd.DataFrame({'客户': ['A']})
    assert gen.check_data_completeness('2025', '12') == []


def test_load_ar_data_groups_cross_year_detail_by_transaction_month(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "公司名称 : 浙江宙恒进出口有限公司 / 2026/01/01  ~ 2026/12/31  / 科目账簿 / 1122(应收账款)"
    headers = [
        "日期-号码", "摘要", "科目名", "科目编码", "相对科目编码名", "相对科目编码",
        "往来单位编码", "往来单位名", "外币借方金额", "外币贷方金额", "借方金额", "贷方金额", "余额",
    ]
    for idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=idx).value = header
    ws.append(["2025/12/31 -1", "年末销售", "应收账款", "1122", "主营业务收入", "6001", "C001", "客户A", 10, None, 70, None, 70])
    ws.append(["2026/01/05 -1", "新年销售", "应收账款", "1122", "主营业务收入", "6001", "C001", "客户A", 20, None, 140, None, 210])
    path = tmp_path / "应收账款2023-2026.xlsx"
    wb.save(path)

    gen = ReportGenerator(str(tmp_path))
    gen._load_ar_data(str(path), path.name)

    assert sorted(gen.data["ar"].keys()) == ["2025-12", "2026-01"]
    assert gen.ar_detail_df is not None
    assert set(gen.ar_detail_df["MonthStr"].unique()) == {"2025-12", "2026-01"}


def test_fill_product_summary_total_uses_weighted_averages():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-11'] = pd.DataFrame([
        {
            'MonthStr': '2025-11',
            'ParsedDate': pd.Timestamp('2025-11-15'),
            '品目编码': '001',
            '品目名': 'A',
            '品目组合1名': '鞋类',
            '数量': 10,
            '合计': 1000,
        },
    ])
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            'MonthStr': '2025-12',
            'ParsedDate': pd.Timestamp('2025-12-15'),
            '品目编码': '001',
            '品目名': 'A',
            '品目组合1名': '鞋类',
            '数量': 10,
            '合计': 1000,
        },
        {
            'MonthStr': '2025-12',
            'ParsedDate': pd.Timestamp('2025-12-15'),
            '品目编码': '002',
            '品目名': 'B',
            '品目组合1名': '鞋类',
            '数量': 1,
            '合计': 100,
        },
    ])
    gen.data['cost']['2025-11'] = pd.DataFrame({
        '品目编码': ['001', '002'],
        '品目名规格': ['A规格', 'B规格'],
        '库存_期初': [100, 50],
        '库存_期初.2': [200, 100],
        '库存_期末': [120, 40],
        '库存_期末.2': [300, 100],
        '单价_减少.1': [50, 10],
    })
    gen.data['cost']['2025-12'] = pd.DataFrame({
        '品目编码': ['001', '002'],
        '品目名规格': ['A规格', 'B规格'],
        '库存_期初': [120, 40],
        '库存_期初.2': [200, 100],
        '库存_期末': [150, 35],
        '库存_期末.2': [400, 100],
        '单价_减少.1': [50, 10],
    })

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按产品汇总_含合计'
    headers = [
        '产品名称',
        '品目编码',
        '年销售数量合计',
        '年销售收入合计',
        '年销售成本合计',
        '年销售利润合计',
        '年初存货金额',
        '年末存货金额',
        '年平均存货',
        '存货周转率',
        '存货周转天数',
        '年销售数量平均',
        '年销售收入平均',
        '年销售成本平均',
        '年销售利润平均',
        '年毛利率平均',
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    ws.cell(row=2, column=2).value = '001'
    ws.cell(row=3, column=2).value = '002'
    ws.cell(row=4, column=1).value = '合计'

    gen._fill_product_summary_total(ws, '2025', '12', 'current')

    # 加权口径（总额分母法）预期值：
    # 年销售数量合计=21，期内月数=2 => 年销售数量平均=10.5
    # 年销售收入合计=2100，年销售成本合计=1010，年销售利润合计=1090
    # 年初存货金额=300，年末存货金额=500 => 年平均存货=400
    # 存货周转率=1010/400=2.525，存货周转天数=365/2.525
    # 年毛利率平均=(1090/2)/(2100/2)=1090/2100
    assert abs(ws.cell(4, 12).value - 10.5) < 1e-12
    assert abs(ws.cell(4, 13).value - 1050) < 1e-12
    assert abs(ws.cell(4, 14).value - 505) < 1e-12
    assert abs(ws.cell(4, 15).value - 545) < 1e-12
    assert abs(ws.cell(4, 9).value - 400) < 1e-12
    assert abs(ws.cell(4, 10).value - 2.525) < 1e-12
    assert abs(ws.cell(4, 11).value - (365 / 2.525)) < 1e-12
    assert abs(ws.cell(4, 16).value - (1090 / 2100)) < 1e-12


def test_fill_product_summary_total_handles_total_marker_and_missing_parsed_date():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            'MonthStr': '2025-12',
            '品目编码': '001',
            '品目名': 'A',
            '品目组合1名': '鞋类',
            '数量': 10,
            '合计': 1000,
        },
    ])
    gen.data['cost']['2025-12'] = pd.DataFrame({
        '品目编码': ['001'],
        '品目名规格': ['A规格'],
        '库存_期初.2': [100],
        '库存_期末.2': [200],
        '单价_减少.1': [20],
    })

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按产品汇总_含合计'
    headers = [
        '产品名称',
        '品目编码',
        '年销售数量合计',
        '年销售收入合计',
        '年销售成本合计',
        '年销售利润合计',
        '年初存货金额',
        '年末存货金额',
        '年平均存货',
        '存货周转率',
        '存货周转天数',
        '年销售数量平均',
        '年销售收入平均',
        '年销售成本平均',
        '年销售利润平均',
        '年毛利率平均',
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    ws.cell(row=2, column=2).value = '001'
    ws.cell(row=3, column=2).value = '合计'
    ws.cell(row=3, column=10).value = 999999  # stale template value
    ws.cell(row=3, column=13).value = 999999  # stale template value

    gen._fill_product_summary_total(ws, '2025', '12', 'current')

    assert abs(ws.cell(row=3, column=13).value - 1000) < 1e-12
    assert abs(ws.cell(row=3, column=10).value - (200 / 150)) < 1e-12


def test_fill_product_summary_total_keeps_total_row_after_inserting_missing_codes():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            'MonthStr': '2025-12',
            '品目编码': '001',
            '品目名': 'A',
            '品目组合1名': '鞋类',
            '数量': 10,
            '合计': 1000,
        },
        {
            'MonthStr': '2025-12',
            '品目编码': '5501',
            '品目名': '5501',
            '品目组合1名': '配件',
            '数量': 1,
            '合计': 100,
        },
    ])
    gen.data['cost']['2025-12'] = pd.DataFrame({
        '品目编码': ['001', '5501'],
        '品目名规格': ['A规格', '5501'],
        '库存_期初.2': [100, 10],
        '库存_期末.2': [200, 20],
        '单价_减少.1': [20, 5],
    })

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按产品汇总_含合计'
    headers = [
        '产品名称',
        '品目编码',
        '年销售数量合计',
        '年销售收入合计',
        '年销售成本合计',
        '年销售利润合计',
        '年初存货金额',
        '年末存货金额',
        '年平均存货',
        '存货周转率',
        '存货周转天数',
        '年销售数量平均',
        '年销售收入平均',
        '年销售成本平均',
        '年销售利润平均',
        '年毛利率平均',
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    # 模板仅有产品001与合计，5501会被插入到合计行之前。
    ws.cell(row=2, column=2).value = '001'
    ws.cell(row=3, column=1).value = '合计'
    ws.cell(row=3, column=3).value = 999999  # stale template total to ensure overwritten

    gen._fill_product_summary_total(ws, '2025', '12', 'current')

    total_row = None
    row_5501 = None
    for r in range(2, ws.max_row + 1):
        first = ws.cell(row=r, column=1).value
        code = ws.cell(row=r, column=2).value
        if first is not None and str(first).strip() == '合计':
            total_row = r
        if code is not None and str(code).strip() == '5501':
            row_5501 = r

    assert total_row is not None
    assert row_5501 is not None
    assert row_5501 < total_row

    # 5501为自身数据，不应被覆盖为总计。
    assert abs(ws.cell(row=row_5501, column=3).value - 1) < 1e-12
    # 总计应为001+5501。
    assert abs(ws.cell(row=total_row, column=3).value - 11) < 1e-12


def test_fill_expense_details_places_anomaly_section_below_main_table():
    gen = ReportGenerator('.')
    gen.data['expense']['2025-12'] = pd.DataFrame([
        {
            '日期': '2025-12-15',
            '科目名': '管理费用-办公费',
            '借方金额': 1200,
            '贷方金额': 0,
            '部门名': '行政',
            '摘要': '办公用品',
        },
    ])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '费用明细'

    gen._fill_expense_details(ws, '2025', '12', 'current')

    assert ws.cell(row=1, column=1).value == '月份'
    # 主表不再并排写异常明细。
    assert ws.cell(row=1, column=9).value is None
    assert "费用异常明细" in wb.sheetnames

    detail_ws = wb["费用异常明细"]
    assert "异常项目明细" in str(detail_ws.cell(row=1, column=1).value)
    assert detail_ws.cell(row=2, column=1).value == '月份'
    # 单条样本通常不会触发异常评分，明细页应给出无可关联提示。
    assert detail_ws.cell(row=3, column=1).value == "无可关联的异常项目明细"


def test_add_chart_expense_detail_prefers_subcategory_dimension():
    gen = ReportGenerator('.')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '费用明细'
    headers = ['月份', '部门', '费用类别', '子科目', '摘要', '金额']
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c).value = h

    rows = [
        ['2025/12', '销售部', '销售费用', '工资', '工资发放', 100],
        ['2025/12', '销售部', '销售费用', '房租', '门店房租', 200],
        ['2025/12', '行政部', '管理费用', '工资', '行政工资', 50],
    ]
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c).value = v

    start_col = ws.max_column + 2
    added = gen._add_chart_expense_detail(ws)
    assert added

    assert ws.cell(row=1, column=start_col).value == '子科目'
    assert ws.cell(row=2, column=start_col).value == '房租'
    assert abs(ws.cell(row=2, column=start_col + 1).value - 200) < 1e-12
    labels = {
        ws.cell(row=2, column=start_col).value,
        ws.cell(row=3, column=start_col).value,
    }
    assert labels == {'房租', '工资'}


def test_ensure_report_charts_rebuilds_sales_inventory_chart_with_fallback_data():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '明细_销售与库存'

    headers = ['品目编码', '产品大类', '销售收入', '期末金额']
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c).value = h

    ws.cell(row=2, column=1).value = '001'
    ws.cell(row=2, column=2).value = '鞋类'
    ws.cell(row=2, column=4).value = 100
    ws.cell(row=3, column=1).value = '002'
    ws.cell(row=3, column=2).value = '鞋类'
    ws.cell(row=3, column=4).value = 50
    ws.cell(row=4, column=1).value = '003'
    ws.cell(row=4, column=2).value = '电器'
    ws.cell(row=4, column=4).value = 80

    # Simulate stale helper area + stale chart from previous generation.
    ws.cell(row=1, column=28).value = '品类'
    ws.cell(row=1, column=29).value = '金额'
    ws.cell(row=2, column=28).value = '旧品类'
    ws.cell(row=2, column=29).value = 999
    ws.cell(row=1, column=40).value = '旧说明'

    stale = BarChart()
    stale.add_data(Reference(ws, min_col=29, max_col=29, min_row=1, max_row=2), titles_from_data=True)
    stale.set_categories(Reference(ws, min_col=28, min_row=2, max_row=2))
    ws.add_chart(stale, 'F2')
    assert len(ws._charts) == 1

    gen._ensure_report_charts(wb)

    assert len(ws._charts) == 1
    chart = ws._charts[0]
    series = chart.series[0]
    val_ref = series.val.numRef.f if series.val is not None and series.val.numRef is not None else None
    cat_ref = None
    if series.cat is not None:
        if series.cat.strRef is not None:
            cat_ref = series.cat.strRef.f
        elif series.cat.numRef is not None:
            cat_ref = series.cat.numRef.f

    assert val_ref == "'明细_销售与库存'!$AC$2:$AC$3"
    assert cat_ref == "'明细_销售与库存'!$AB$2:$AB$3"

    # Sales revenue column is empty, chart should fallback to ending inventory amount.
    assert ws.cell(row=2, column=28).value == '鞋类'
    assert abs(ws.cell(row=2, column=29).value - 150) < 1e-12
    assert ws.cell(row=3, column=28).value == '电器'
    assert abs(ws.cell(row=3, column=29).value - 80) < 1e-12


def test_add_pareto_chart_uses_header_series_titles():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '客户贡献与回款'

    ws.cell(row=1, column=1).value = '客户'
    ws.cell(row=1, column=2).value = '销售收入'
    ws.cell(row=1, column=3).value = '累计占比'
    ws.cell(row=2, column=1).value = 'A客户'
    ws.cell(row=2, column=2).value = 100
    ws.cell(row=2, column=3).value = 0.5
    ws.cell(row=3, column=1).value = 'B客户'
    ws.cell(row=3, column=2).value = 60
    ws.cell(row=3, column=3).value = 0.8
    ws.cell(row=4, column=1).value = 'C客户'
    ws.cell(row=4, column=2).value = 40
    ws.cell(row=4, column=3).value = 1.0

    gen._add_pareto_chart(ws, 1, 2, 3, 1, 2, 4, '客户收入集中度 (Pareto)', 'E2')
    chart = ws._charts[-1]

    title_tokens = []
    for sub_chart in [chart] + list(getattr(chart, '_charts', [])):
        for series in sub_chart.series:
            token = _series_title_token(series)
            if token:
                title_tokens.append(token.replace('$', ''))

    assert any(token.endswith('!B1') for token in title_tokens)
    assert any(token.endswith('!C1') for token in title_tokens)
    assert all(token not in {'系列1', '系列2', 'Series1', 'Series2'} for token in title_tokens)


def test_add_scatter_chart_uses_y_header_as_series_title():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '销售人效分析'

    ws.cell(row=1, column=1).value = '人均收入'
    ws.cell(row=1, column=2).value = '利润率'
    ws.cell(row=2, column=1).value = 120
    ws.cell(row=2, column=2).value = 0.15
    ws.cell(row=3, column=1).value = 95
    ws.cell(row=3, column=2).value = 0.1

    gen._add_scatter_chart(ws, 1, 2, 1, 2, 3, '散点图', 'E2', x_title='人均收入', y_title='利润率')
    chart = ws._charts[-1]
    series_title = _series_title_token(chart.series[0])

    assert series_title == '利润率'
    assert series_title not in {'系列1', 'Series1'}


def test_add_doughnut_chart_uses_header_series_title():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '费用结构'

    ws.cell(row=1, column=1).value = '费用类型'
    ws.cell(row=1, column=2).value = '金额'
    ws.cell(row=2, column=1).value = '管理费用'
    ws.cell(row=2, column=2).value = 100
    ws.cell(row=3, column=1).value = '销售费用'
    ws.cell(row=3, column=2).value = 80

    gen._add_doughnut_chart(ws, 1, 2, 1, 2, 3, '费用构成', 'E2')
    chart = ws._charts[-1]
    series_title = _series_title_token(chart.series[0])

    assert series_title is not None
    assert series_title.replace('$', '').endswith('!B1')


def test_write_chart_note_normalizes_oversized_row_height():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.row_dimensions[5].height = 24.0
    ws.row_dimensions[6].height = 199.5
    ws.row_dimensions[7].height = 24.0

    gen._write_chart_note(ws, 1, 6, "图表说明：测试")

    assert abs(ws.row_dimensions[6].height - 24.0) < 1e-12


def test_append_chart_notes_below_keeps_note_row_height_and_disables_wrap():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '经营指标'

    ws.cell(row=1, column=1).value = '月份'
    ws.cell(row=1, column=2).value = '收入'
    ws.cell(row=2, column=1).value = '2025-11'
    ws.cell(row=2, column=2).value = 100
    ws.cell(row=3, column=1).value = '2025-12'
    ws.cell(row=3, column=2).value = 120

    chart = BarChart()
    chart.height = 10
    chart.width = 10
    data_ref = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=3)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=3)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    ws.add_chart(chart, 'E2')

    _, _, _, row_end = gen._extract_chart_anchor_bbox(ws, chart)
    note_row = row_end + 2
    ws.row_dimensions[note_row - 1].height = 24.0
    ws.row_dimensions[note_row].height = 213.75
    ws.row_dimensions[note_row + 1].height = 24.0

    gen._append_chart_notes_below(wb, '2025', '12')

    note_cells = [
        cell for cell in ws._cells.values()
        if isinstance(cell.value, str) and '图表说明（2025年12月）' in cell.value
    ]
    assert note_cells
    note_cell = note_cells[0]
    assert note_cell.row == note_row
    assert abs(ws.row_dimensions[note_row].height - 24.0) < 1e-12
    assert note_cell.alignment is not None
    assert note_cell.alignment.wrap_text is False


def test_append_chart_notes_below_deduplicates_adjacent_same_notes():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '经营指标'

    ws.cell(row=1, column=1).value = '月份'
    ws.cell(row=1, column=2).value = '收入'
    ws.cell(row=2, column=1).value = '2025-11'
    ws.cell(row=2, column=2).value = 100
    ws.cell(row=3, column=1).value = '2025-12'
    ws.cell(row=3, column=2).value = 120

    for anchor in ('E2', 'E3'):
        chart = BarChart()
        chart.height = 10
        chart.width = 6
        data_ref = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=3)
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=3)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, anchor)

    gen._append_chart_notes_below(wb, '2025', '12')

    notes = [
        cell for cell in ws._cells.values()
        if isinstance(cell.value, str) and '图表说明（2025年12月）' in cell.value
    ]
    assert len(notes) == 1


def test_data_quality_prefers_parsed_date_for_sales():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            '日期-号码': 'NO_DATE_TOKEN',
            'ParsedDate': pd.Timestamp('2025-12-15'),
            '品目编码': '001',
            '数量': 2,
            '销售金额合计': 100,
            '销售出库供应价合计': 60,
            '往来单位名': '客户A',
            '销售订单号': 'SO001',
        },
        {
            '日期-号码': 'INVALID',
            'ParsedDate': pd.Timestamp('2025-12-20'),
            '品目编码': '002',
            '数量': 3,
            '销售金额合计': 120,
            '销售出库供应价合计': 70,
            '往来单位名': '客户B',
            '销售订单号': 'SO002',
        },
    ])

    gen._run_data_quality_checks()
    sales_date_fail = [
        item for item in gen.data_quality_issues
        if item.get('category') == 'sales' and item.get('issue_type') == '日期解析失败'
    ]
    assert not sales_date_fail


def test_data_quality_skips_month_mismatch_for_multi_period_ledger():
    gen = ReportGenerator('.')
    gen.data['ar']['2026-12'] = pd.DataFrame([
        {'日期': '2025-01-10', '往来单位名': 'A', '借方金额': 1, '贷方金额': 0},
        {'日期': '2025-02-10', '往来单位名': 'B', '借方金额': 1, '贷方金额': 0},
        {'日期': '2025-03-10', '往来单位名': 'C', '借方金额': 1, '贷方金额': 0},
        {'日期': '2025-04-10', '往来单位名': 'D', '借方金额': 1, '贷方金额': 0},
    ])

    gen._run_data_quality_checks()
    ar_month_mismatch = [
        item for item in gen.data_quality_issues
        if item.get('category') == 'ar' and item.get('issue_type') == '月份不匹配'
    ]
    assert not ar_month_mismatch


def test_data_quality_sales_duplicate_order_is_info_not_warn():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            '日期': '2025-12-10',
            '品目编码': '001',
            '数量': 2,
            '销售金额合计': 100,
            '销售出库供应价合计': 60,
            '往来单位名': '客户A',
            '销售订单号': 'SO001',
        },
        {
            '日期': '2025-12-11',
            '品目编码': '002',
            '数量': 3,
            '销售金额合计': 120,
            '销售出库供应价合计': 70,
            '往来单位名': '客户A',
            '销售订单号': 'SO001',
        },
    ])

    gen._run_data_quality_checks()
    dup_items = [
        item for item in gen.data_quality_issues
        if item.get('category') == 'sales' and item.get('issue_type') == '单号重复'
    ]
    assert dup_items
    assert all((item.get('severity') or '').upper() == 'INFO' for item in dup_items)


def test_data_quality_summary_for_scope_excludes_out_of_scope_periods():
    gen = ReportGenerator('.')
    gen.data_quality_issues = [
        {'severity': 'ERROR', 'category': 'sales', 'period': '2026-02', 'issue_type': '客户/单位缺失', 'detail': 'x'},
        {'severity': 'WARN', 'category': 'sales', 'period': '2025-12', 'issue_type': '单价异常', 'detail': 'x'},
        {'severity': 'INFO', 'category': 'expense', 'period': '2025-11', 'issue_type': '金额/数量异常值', 'detail': 'x'},
    ]

    summary = gen._get_data_quality_summary_for_scope('2025', '12', 'current')
    assert summary['ERROR'] == 0
    assert summary['WARN'] == 1
    assert summary['INFO'] == 1
    assert summary['TOTAL'] == 2


def test_dashboard_template_formulas_use_prev_month_lookup_not_match_minus_one():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '仪表盘'

    # Trigger template-formula mode
    ws['A5'].value = '=OLD_FORMULA'
    ws['A7'].value = '=OLD_DELTA'
    ws['B3'].value = '2025/12'
    ws['A4'].value = '主营业务收入（元）'
    ws['E4'].value = '净利润（元）'
    ws['I4'].value = '净利润率'
    ws['M4'].value = '成本率'
    ws['Q4'].value = '营业利润（元）'
    ws['A1'].value = '较上一年：示例'

    gen._update_dashboard(wb, {}, '2025', '12')

    prev_lookup_token = 'MATCH(TEXT(DATE(LEFT($B$3,4),RIGHT($B$3,2),1)-1,"yyyy/mm")'

    assert prev_lookup_token in ws['A7'].value
    assert 'MATCH($B$3' in ws['A5'].value
    assert 'MATCH($B$3' in ws['E5'].value
    assert "'利润表'" in ws['E7'].value
    assert "'经营指标'!$G:$G" in ws['Q5'].value
    assert 'MATCH($B$3' in ws['Q5'].value
    assert prev_lookup_token in ws['I7'].value
    assert '*100' in ws['I7'].value
    assert '*100' in ws['M7'].value
    assert ws['A1'].value == '较上月：示例'


def test_ensure_month_columns_simple_backfills_intermediate_months():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '利润表'
    ws.cell(row=1, column=1).value = '指标'
    ws.cell(row=1, column=2).value = '全年汇总'
    ws.cell(row=1, column=3).value = '2025/11'
    ws.cell(row=1, column=4).value = '2025/10'
    ws.cell(row=1, column=5).value = '2025/09'

    changed = gen._ensure_month_columns_simple(ws, '2026', '2', header_row=1)

    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    assert changed is True
    assert headers[:7] == ['指标', '全年汇总', '2026/02', '2026/01', '2025/12', '2025/11', '2025/10']


def test_ensure_month_rows_simple_backfills_intermediate_months():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '本量利分析'
    ws.cell(row=1, column=1).value = '月份'
    ws.cell(row=1, column=2).value = '部门'
    ws.cell(row=2, column=1).value = '2025/11'
    ws.cell(row=2, column=2).value = '合计'
    ws.cell(row=3, column=1).value = '2025/10'
    ws.cell(row=3, column=2).value = '合计'

    changed = gen._ensure_month_rows_simple(ws, '2026', '2', total_label='')
    gen._reorder_month_rows_desc(ws)

    labels = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    assert changed is True
    assert labels[:4] == ['2026/02', '2026/01', '2025/12', '2025/11']


def test_ensure_month_columns_grouped_by_suffix_backfills_intermediate_months():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按品类汇总(按月)'
    headers = [
        '产品大类',
        '2025/11_毛利润',
        '2025/10_毛利润',
        '2025/11_销售成本',
        '2025/10_销售成本',
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    changed = gen._ensure_month_columns_grouped_by_suffix(ws, '2026', '2')

    out = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    assert changed is True
    assert '2026/02_毛利润' in out
    assert '2026/01_毛利润' in out
    assert '2025/12_毛利润' in out
    assert '2026/02_销售成本' in out
    assert '2026/01_销售成本' in out
    assert '2025/12_销售成本' in out


def test_load_ar_data_appends_multiple_files(tmp_path):
    gen = ReportGenerator('.')

    path1 = tmp_path / '2023-2025应收账款.xlsx'
    path2 = tmp_path / '2026应收账款.xlsx'

    df1 = pd.DataFrame({
        '日期': ['2025-12-15'],
        '往来单位名': ['客户A'],
        '借方金额': [100],
        '贷方金额': [0],
    })
    df2 = pd.DataFrame({
        '日期': ['2026-01-15'],
        '往来单位名': ['客户A'],
        '借方金额': [50],
        '贷方金额': [0],
    })

    with pd.ExcelWriter(path1) as writer:
        pd.DataFrame([['dummy']]).to_excel(writer, index=False, header=False)
        df1.to_excel(writer, index=False, startrow=1)
    with pd.ExcelWriter(path2) as writer:
        pd.DataFrame([['dummy']]).to_excel(writer, index=False, header=False)
        df2.to_excel(writer, index=False, startrow=1)

    gen._load_ar_data(str(path1), path1.name)
    gen._load_ar_data(str(path2), path2.name)

    assert set(gen.ar_detail_df['MonthStr'].unique()) == {'2025-12', '2026-01'}
    assert '2025-12' in gen.data['ar']
    assert '2026-01' in gen.data['ar']


def test_apply_current_scope_visibility_hides_out_of_scope_months():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '利润表'
    ws.cell(row=1, column=1).value = '指标'
    ws.cell(row=1, column=2).value = '2026/02'
    ws.cell(row=1, column=3).value = '2026/01'
    ws.cell(row=1, column=4).value = '2025/12'

    gen._apply_current_scope_visibility(wb, '2026', '2', 'current')

    assert ws.column_dimensions['B'].hidden is False
    assert ws.column_dimensions['C'].hidden is False
    assert ws.column_dimensions['D'].hidden is True


def test_hide_leading_blank_rows_hides_gap_before_first_content():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1).value = '月份'
    ws.cell(row=5, column=1).value = '2026/02'

    gen._hide_leading_blank_rows(ws, header_row=1)

    assert ws.row_dimensions[2].hidden is True
    assert ws.row_dimensions[3].hidden is True
    assert ws.row_dimensions[4].hidden is True


def test_hide_rows_before_first_month_hides_gap_before_month_data():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=1, column=1).value = '月份'
    ws.cell(row=2, column=2).value = '=IF(1=0,\"\",\"\")'
    ws.cell(row=5, column=1).value = '2026/02'

    gen._hide_rows_before_first_month(ws, header_keyword='月份', month_col=1)

    assert ws.row_dimensions[2].hidden is True
    assert ws.row_dimensions[3].hidden is True
    assert ws.row_dimensions[4].hidden is True


def test_validate_report_accepts_equivalent_b3_dropdown_formula():
    gen = ReportGenerator('.')

    wb = openpyxl.Workbook()

    ws_profit = wb.active
    ws_profit.title = '利润表'
    ws_profit['A1'].value = '项目'
    ws_profit['B1'].value = '2025/12'

    ws_asset = wb.create_sheet('资产负债表')
    ws_asset['A1'].value = '项目'
    ws_asset['B1'].value = '2025/12'

    ws_metrics = wb.create_sheet('经营指标')
    ws_metrics['A1'].value = '月份'
    ws_metrics['A2'].value = '2025/12'

    ws_dashboard = wb.create_sheet('仪表盘')
    ws_dashboard['B3'].value = '2025/12'
    dv = DataValidation(type='list', formula1='经营指标!$A$2:$A$2', allow_blank=True)
    ws_dashboard.add_data_validation(dv)
    dv.add('B3')

    tmp_path = '__tmp_validate_b3_equivalent.xlsx'
    wb.save(tmp_path)
    try:
        issues = gen.validate_report_file(tmp_path, '2025', '12', 'current')
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    assert not any('B3下拉范围未匹配' in str(item.get('message') or '') for item in issues)


def test_update_dashboard_controls_updates_unquoted_metric_chart_ranges():
    gen = ReportGenerator('.')

    wb = openpyxl.Workbook()
    ws_metric = wb.active
    ws_metric.title = '经营指标'
    ws_metric['A1'].value = '月份'
    ws_metric['A13'].value = '2026/02'
    ws_metric['A14'].value = '2026/01'
    ws_metric['C1'].value = '主营业务收入'
    ws_metric['C13'].value = 100
    ws_metric['C14'].value = 90

    ws_dash = wb.create_sheet('仪表盘')
    chart = openpyxl.chart.LineChart()
    data = Reference(ws_metric, min_col=3, min_row=1, max_row=12)
    cats = Reference(ws_metric, min_col=1, min_row=2, max_row=12)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws_dash.add_chart(chart, 'A1')

    gen._update_dashboard_controls(wb, '2026', '2', 'current')

    series = ws_dash._charts[0].series[0]
    assert series.val.numRef.f == '经营指标!$C$13:$C$14'
    assert series.cat.numRef.f == '经营指标!$A$13:$A$14'


def test_trim_chart_data_ranges_shrinks_to_active_rows():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '经营指标'
    ws['A1'].value = '月份'
    ws['C1'].value = '主营业务收入'
    ws['A2'].value = None
    ws['A3'].value = '2026/02'
    ws['A4'].value = '2026/01'
    ws['A5'].value = None
    ws['C3'].value = 100
    ws['C4'].value = 90

    chart = openpyxl.chart.LineChart()
    data = Reference(ws, min_col=3, min_row=1, max_row=5)
    cats = Reference(ws, min_col=1, min_row=2, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, 'E1')

    changed = gen._trim_single_chart_data_range(wb, ws._charts[0])

    series = ws._charts[0].series[0]
    assert changed is True
    assert series.val.numRef.f == "'经营指标'!$C$3:$C$4" or series.val.numRef.f == '经营指标!$C$3:$C$4'
    assert series.cat.numRef.f == "'经营指标'!$A$3:$A$4" or series.cat.numRef.f == '经营指标!$A$3:$A$4'


def test_management_metrics_sheet_includes_yoy_and_mom_columns():
    gen = ReportGenerator('.')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '经营指标'
    headers = ['月份', '部门', '主营业务收入', '主营业务成本', '销售费用', '管理费用', '营业利润']
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c).value = h
    ws.cell(row=2, column=1).value = '2025/12'
    ws.cell(row=3, column=1).value = '2025/11'

    scoped = {
        '2025-11': {
            'revenue': 100,
            'cost': 60,
            'sales_expense': 10,
            'admin_expense': 5,
            'operating_profit': 25,
        },
        '2025-12': {
            'revenue': 120,
            'cost': 70,
            'sales_expense': 12,
            'admin_expense': 6,
            'operating_profit': 32,
        },
    }
    all_metrics = {
        **scoped,
        '2024-12': {
            'revenue': 80,
            'cost': 50,
            'sales_expense': 8,
            'admin_expense': 4,
            'operating_profit': 18,
        },
    }

    gen._update_management_metrics_sheet(
        ws,
        scoped,
        '2025',
        '12',
        'current',
        metrics_by_month_all=all_metrics,
    )

    header_map = {str(ws.cell(row=1, column=c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value}
    row_2025_12 = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value).strip() == '2025/12':
            row_2025_12 = r
            break
    assert row_2025_12 is not None

    assert abs(ws.cell(row_2025_12, column=header_map['主营业务收入_同比增量']).value - 40) < 1e-12
    assert abs(ws.cell(row_2025_12, column=header_map['主营业务收入_同比增速']).value - 0.5) < 1e-12
    assert abs(ws.cell(row_2025_12, column=header_map['主营业务收入_环比增量']).value - 20) < 1e-12
    assert abs(ws.cell(row_2025_12, column=header_map['主营业务收入_环比增速']).value - 0.2) < 1e-12


def test_category_month_sheet_labels_show_revenue_share_and_remain_idempotent():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            'MonthStr': '2025-12',
            'ParsedDate': pd.Timestamp('2025-12-15'),
            '品目编码': '001',
            '品目组合1名': '电器类',
            '数量': 3,
            '合计': 300,
        },
        {
            'MonthStr': '2025-12',
            'ParsedDate': pd.Timestamp('2025-12-20'),
            '品目编码': '002',
            '品目组合1名': '鞋类',
            '数量': 1,
            '合计': 100,
        },
    ])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按品类汇总(按月)'
    headers = ['产品大类', '年销售数量', '年销售收入', '年销售成本', '年毛利润', '2025/12_毛利润']
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    ws.cell(row=2, column=1).value = '电器类'
    ws.cell(row=3, column=1).value = '鞋类'
    ws.cell(row=4, column=1).value = '合计'

    gen._update_category_month_sheet(ws, '2025', '12', 'current')

    assert ws.cell(row=2, column=1).value == '电器类'
    assert ws.cell(row=3, column=1).value == '电器类占比'
    assert ws.cell(row=4, column=1).value == '鞋类'
    assert ws.cell(row=5, column=1).value == '鞋类占比'
    assert abs(ws.cell(row=3, column=3).value - 0.75) < 1e-12
    assert abs(ws.cell(row=5, column=3).value - 0.25) < 1e-12

    # 再次执行不应叠加“占比”文本，且仍可正确匹配并回写。
    gen._update_category_month_sheet(ws, '2025', '12', 'current')
    assert ws.cell(row=2, column=1).value == '电器类'
    assert ws.cell(row=3, column=1).value == '电器类占比'
    assert ws.cell(row=4, column=1).value == '鞋类'
    assert ws.cell(row=5, column=1).value == '鞋类占比'


def test_category_contribution_is_merged_and_old_sheet_redirects():
    gen = ReportGenerator('.')
    gen.data['sales']['2025-12'] = pd.DataFrame([
        {
            'MonthStr': '2025-12',
            '品目编码': '001',
            '品目名': 'A',
            '品目组合1名': '鞋类',
            '数量': 2,
            '合计': 100,
        },
        {
            'MonthStr': '2025-12',
            '品目编码': '002',
            '品目名': 'B',
            '品目组合1名': '电器类',
            '数量': 1,
            '合计': 80,
        },
    ])
    gen.data['cost']['2025-12'] = pd.DataFrame({
        '品目编码': ['001', '002'],
        '单价_减少.1': [30, 20],
    })

    wb = openpyxl.Workbook()
    ws_month = wb.active
    ws_month.title = '按品类汇总(按月)'
    ws_month.cell(row=1, column=1).value = '产品大类'
    wb.create_sheet('品类贡献毛利')

    gen._update_category_contribution_sheet(wb, '2025', '12', 'current')

    merge_col = None
    for c in range(1, ws_month.max_column + 1):
        if str(ws_month.cell(row=1, column=c).value).strip() == '品类贡献分析(合并视图)':
            merge_col = c
            break
    assert merge_col is not None
    assert ws_month.cell(row=4, column=merge_col).value == '品类'


def test_delete_merged_sheets_keeps_expense_details():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    wb.active.title = '经营指标'
    for name in [
        '同比经营分析',
        '环比经营分析',
        '品类贡献毛利',
        '费用明细环比分析',
        '费用结构与弹性',
        '异常预警',
        '年度费用异常Top',
        '费用明细',
    ]:
        wb.create_sheet(name)

    gen._delete_sheets_if_exist(
        wb,
        [
            '同比经营分析',
            '环比经营分析',
            '品类贡献毛利',
            '费用明细环比分析',
            '费用结构与弹性',
            '异常预警',
            '年度费用异常Top',
        ],
    )

    for removed in [
        '同比经营分析',
        '环比经营分析',
        '品类贡献毛利',
        '费用明细环比分析',
        '费用结构与弹性',
        '异常预警',
        '年度费用异常Top',
    ]:
        assert removed not in wb.sheetnames
    assert '费用明细' in wb.sheetnames


def test_expense_diagnostic_center_generates_and_links_details():
    gen = ReportGenerator('.')
    gen.data['expense']['2025-11'] = pd.DataFrame([
        {
            '日期': '2025-11-15',
            '科目名': '管理费用-办公费',
            '借方金额': 100,
            '贷方金额': 0,
            '部门名': '行政',
            '摘要': '办公用品',
        },
    ])
    gen.data['expense']['2025-12'] = pd.DataFrame([
        {
            '日期': '2025-12-15',
            '科目名': '管理费用-办公费',
            '借方金额': 12000,
            '贷方金额': 0,
            '部门名': '行政',
            '摘要': '年末集中采购',
        },
    ])

    metrics = {
        '2025-11': {
            'revenue': 100000,
            'cost': 70000,
            'sales_expense': 3000,
            'admin_expense': 2000,
            'operating_profit': 15000,
            'inventory_start': 50000,
            'inventory_end': 52000,
            'cost_rate': 0.7,
            'ar_balance': 20000,
        },
        '2025-12': {
            'revenue': 105000,
            'cost': 73000,
            'sales_expense': 3200,
            'admin_expense': 12500,
            'operating_profit': 12000,
            'inventory_start': 52000,
            'inventory_end': 58000,
            'cost_rate': 73000 / 105000,
            'ar_balance': 26000,
        },
    }

    wb = openpyxl.Workbook()
    wb.active.title = '费用明细环比分析'
    gen._update_expense_diagnostic_center(wb, metrics, '2025', '12', 'current', anomaly_top_n=20, matrix_top_n=20, detail_lines_per_key=2)

    assert '费用分析' in wb.sheetnames
    ws = wb['费用分析']
    assert '费用分析' in str(ws.cell(row=1, column=1).value)

    has_anomaly_section = any(
        str(ws.cell(row=r, column=1).value).startswith('C. 异常Top')
        for r in range(1, min(ws.max_row, 200) + 1)
        if ws.cell(row=r, column=1).value is not None
    )
    assert has_anomaly_section

    has_internal_link = False
    for r in range(1, min(ws.max_row, 400) + 1):
        for c in range(1, min(ws.max_column, 40) + 1):
            link = ws.cell(row=r, column=c).hyperlink
            if link and link.location and "费用分析'!A" in link.location:
                has_internal_link = True
                break
        if has_internal_link:
            break
    assert has_internal_link


def test_drilldown_links_point_to_expense_diagnostic_center():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    wb.active.title = '经营指标'
    wb.create_sheet('费用对比')
    wb.create_sheet('费用分析')
    wb.create_sheet('利润表')

    gen._add_drilldown_links(wb)

    ws_metric = wb['经营指标']
    metric_links = [
        ws_metric.cell(row=r, column=ws_metric.max_column).hyperlink.location
        for r in range(1, ws_metric.max_row + 1)
        if ws_metric.cell(row=r, column=ws_metric.max_column).hyperlink is not None
    ]
    assert "'费用分析'!A1" in metric_links

    ws_exp = wb['费用对比']
    exp_links = [
        ws_exp.cell(row=r, column=ws_exp.max_column).hyperlink.location
        for r in range(1, ws_exp.max_row + 1)
        if ws_exp.cell(row=r, column=ws_exp.max_column).hyperlink is not None
    ]
    assert exp_links == ["'费用分析'!A1"]


def test_fill_profit_sheet_refreshes_annual_total_column():
    gen = ReportGenerator('.')
    gen.data['profit']['2025-11'] = pd.DataFrame([
        {'项目': '四、净利润', '2025/11': 100},
    ])
    gen.data['profit']['2025-12'] = pd.DataFrame([
        {'项目': '四、净利润', '2025/12': 200},
    ])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '利润表'
    ws.cell(row=1, column=1).value = '指标'
    ws.cell(row=1, column=2).value = '全年汇总'
    ws.cell(row=1, column=3).value = '2025/12'
    ws.cell(row=1, column=4).value = '2025/11'
    ws.cell(row=2, column=1).value = '四、净利润'
    ws.cell(row=2, column=2).value = 999999  # stale template value

    gen._fill_profit_sheet(ws, '2025', '12', 'current')

    assert ws.cell(row=2, column=3).value == 200
    assert ws.cell(row=2, column=4).value == 100
    assert ws.cell(row=2, column=2).value == 300


def test_profit_sheet_highlights_large_expense_mom_and_links_to_expense_details():
    gen = ReportGenerator('.')
    gen.data['expense']['2025-11'] = pd.DataFrame([
        {
            'MonthStr': '2025-11',
            '科目名': '管理费用-房租',
            '借方金额': 10000,
            '贷方金额': 0,
            '部门名': '行政',
            '摘要': '11月房租',
        },
    ])
    gen.data['expense']['2025-12'] = pd.DataFrame([
        {
            'MonthStr': '2025-12',
            '科目名': '管理费用-房租',
            '借方金额': 30000,
            '贷方金额': 0,
            '部门名': '行政',
            '摘要': '12月房租',
        },
    ])

    wb = openpyxl.Workbook()
    ws_profit = wb.active
    ws_profit.title = '利润表'
    ws_profit.cell(row=1, column=1).value = '项目'
    ws_profit.cell(row=1, column=2).value = '2025/11'
    ws_profit.cell(row=1, column=3).value = '2025/12'
    ws_profit.cell(row=2, column=1).value = '管理费用-房租'
    ws_profit.cell(row=2, column=2).value = 10000
    ws_profit.cell(row=2, column=3).value = 30000
    ws_profit.cell(row=3, column=1).value = '管理费用-办公费'
    ws_profit.cell(row=3, column=2).value = 5000
    ws_profit.cell(row=3, column=3).value = 5200

    ws_expense = wb.create_sheet('费用明细')
    headers = ["月份", "部门", "费用类别", "子科目", "摘要", "金额", "异常标签", "月份键"]
    for c, h in enumerate(headers, start=1):
        ws_expense.cell(row=1, column=c).value = h
    ws_expense.cell(row=2, column=1).value = '2025/11'
    ws_expense.cell(row=2, column=2).value = '行政'
    ws_expense.cell(row=2, column=3).value = '管理费用'
    ws_expense.cell(row=2, column=4).value = '房租'
    ws_expense.cell(row=2, column=8).value = '2025-11'
    ws_expense.cell(row=3, column=1).value = '2025/12'
    ws_expense.cell(row=3, column=2).value = '行政'
    ws_expense.cell(row=3, column=3).value = '管理费用'
    ws_expense.cell(row=3, column=4).value = '房租'
    ws_expense.cell(row=3, column=8).value = '2025-12'

    gen._highlight_profit_expense_anomalies(wb, '2025', '12', 'current')

    flagged = ws_profit.cell(row=2, column=3)
    assert flagged.hyperlink is not None
    assert "'费用明细'!A3" in str(flagged.hyperlink.location)
    assert flagged.font is not None
    assert flagged.font.color is not None
    assert (flagged.font.color.rgb or '').upper() in ('00FF0000', 'FFFF0000')

    # Non-anomalous row should remain without hyperlink.
    assert ws_profit.cell(row=3, column=3).hyperlink is None


def test_update_directory_sheet_realigns_links_and_clears_missing_targets():
    gen = ReportGenerator('.')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '目录'
    ws.cell(row=1, column=1).value = '目标_预算'
    ws.cell(row=1, column=1).hyperlink = "#'明细_销售与库存'!A1"  # stale wrong link
    ws.cell(row=2, column=1).value = '同比经营分析'  # missing sheet
    ws.cell(row=2, column=1).hyperlink = "#'同比经营分析'!A1"
    wb.create_sheet('目标_预算')
    wb.create_sheet('明细_销售与库存')

    gen._update_directory_sheet(wb)

    assert ws.cell(row=1, column=1).hyperlink is not None
    assert ws.cell(row=1, column=1).hyperlink.location == "'目标_预算'!A1"
    assert ws.cell(row=2, column=1).hyperlink is None


if __name__ == '__main__':
    test_fill_product_summary_aggregates_rows()
    test_list_available_months_uses_core_intersection_when_loaded()
    test_check_data_completeness_includes_sales_and_ar()
    test_fill_product_summary_total_uses_weighted_averages()
    test_fill_product_summary_total_handles_total_marker_and_missing_parsed_date()
    test_fill_product_summary_total_keeps_total_row_after_inserting_missing_codes()
    test_fill_expense_details_places_anomaly_section_below_main_table()
    test_add_chart_expense_detail_prefers_subcategory_dimension()
    test_ensure_report_charts_rebuilds_sales_inventory_chart_with_fallback_data()
    test_add_pareto_chart_uses_header_series_titles()
    test_add_scatter_chart_uses_y_header_as_series_title()
    test_add_doughnut_chart_uses_header_series_title()
    test_write_chart_note_normalizes_oversized_row_height()
    test_append_chart_notes_below_keeps_note_row_height_and_disables_wrap()
    test_append_chart_notes_below_deduplicates_adjacent_same_notes()
    test_data_quality_prefers_parsed_date_for_sales()
    test_data_quality_skips_month_mismatch_for_multi_period_ledger()
    test_data_quality_sales_duplicate_order_is_info_not_warn()
    test_data_quality_summary_for_scope_excludes_out_of_scope_periods()
    test_dashboard_template_formulas_use_prev_month_lookup_not_match_minus_one()
    test_validate_report_accepts_equivalent_b3_dropdown_formula()
    test_management_metrics_sheet_includes_yoy_and_mom_columns()
    test_category_month_sheet_labels_show_revenue_share_and_remain_idempotent()
    test_category_contribution_is_merged_and_old_sheet_redirects()
    test_delete_merged_sheets_keeps_expense_details()
    test_expense_diagnostic_center_generates_and_links_details()
    test_drilldown_links_point_to_expense_diagnostic_center()
    test_fill_profit_sheet_refreshes_annual_total_column()
    test_profit_sheet_highlights_large_expense_mom_and_links_to_expense_details()
    test_update_directory_sheet_realigns_links_and_clears_missing_targets()
    print('PASS: test_report_generator_repairs')
