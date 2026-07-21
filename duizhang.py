import pandas as pd
import numpy as np

def financial_reconciliation(file_path):
    print("=" * 50)
    print("【财务对账智能诊断系统启动】")
    print("=" * 50)
    
    # 读取 Excel 文件
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"[-] 读取文件失败: {e}")
        return

    print(f"[+] 成功读取文件，总计加载 {len(df)} 行数据。")
    print(f"[+] 当前表格包含的列名有：{list(df.columns)}")
    
    # ----------------------------------------------------
    # TODO: 请把下面这两个名字改成你 Excel 表格里实际的列名！
    # ----------------------------------------------------
    col_a = '请修改为第一列的名字'   # 例如 '应收金额'
    col_b = '请修改为第二列的名字'   # 例如 '实收金额'
    
    if col_a not in df.columns or col_b not in df.columns:
        print(f"\n[-] 提示：代码里的列名不对，请把代码里的 col_a 和 col_b 改成你表格里的实际列名。")
        return

    # 格式清洗，防止文本、空格导致对不上
    df[col_a] = pd.to_numeric(df[col_a], errors='coerce').fillna(0)
    df[col_b] = pd.to_numeric(df[col_b], errors='coerce').fillna(0)

    # 计算总和
    sum_a = df[col_a].sum()
    sum_b = df[col_b].sum()
    total_diff = sum_a - sum_b
    
    print("-" * 50)
    print(f"【总体对账结果】")
    print(f"• {col_a} 总和: {sum_a:,.2f}")
    print(f"• {col_b} 总和: {sum_b:,.2f}")
    print(f"• 总体净差异:   {total_diff:,.2f}")
    
    if abs(total_diff) < 0.001:
        print("[√] 两列总和完全相等！")
    else:
        print("[x] 总和不平，正在筛选具体错账...")
        is_match = np.isclose(df[col_a], df[col_b], atol=0.01)
        error_rows = df[~is_match].copy()
        error_rows['单行差值'] = error_rows[col_a] - error_rows[col_b]
        
        print(f"• 共有 {len(error_rows)} 行数据存在金额不符。")
        print(error_rows[[col_a, col_b, '单行差值']].head(20))
        
        # 导出结果
        error_rows.to_excel('对账异常明细结果.xlsx', index=False)
        print(f"\n[√] 异常明细已导出到文件: 【对账异常明细结果.xlsx】")
    print("=" * 50)