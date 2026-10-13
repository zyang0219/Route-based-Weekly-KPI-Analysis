import re
import requests
import pandas as pd

# ========== 配置 ==========
APP_ID = "cli_aaaa63af823a9ce3"
APP_SECRET = "yljUK3LoJCy4xznQRq9LMbuyd6fBnfv2"
WIKI_NODE_TOKEN = "BaYRwuHQqivQNskYqA6cGihLnhg"  # 从Wiki链接中截取的节点token


# ========== 飞书API相关函数 ==========

def get_tenant_access_token():
    """获取飞书租户访问令牌"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    resp = requests.post(url, json={"app_id": APP_ID, "app_secret": APP_SECRET})
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取token失败: {data}")
    return data["tenant_access_token"]


def get_spreadsheet_token(access_token, wiki_node_token):
    """把Wiki节点解析为真正的电子表格token"""
    url = "https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"token": wiki_node_token}
    resp = requests.get(url, headers=headers, params=params)
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"解析Wiki节点失败: {data}")
    node = data["data"]["node"]
    print("文档类型:", node["obj_type"])  # 应该是 sheet
    return node["obj_token"]


def get_sheet_list(access_token, spreadsheet_token):
    """获取电子表格下所有Sheet列表"""
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取Sheet列表失败: {data}")
    return data["data"]["sheets"]


def get_latest_week_sheet(sheets):
    """从Sheet列表中找出形如 Wxx 的sheet，返回周数最大的那个"""
    week_sheets = []
    for s in sheets:
        title = s["title"]
        match = re.match(r"^W(\d+)$", title)
        if match:
            week_sheets.append((int(match.group(1)), s))
    if not week_sheets:
        raise Exception("未找到形如 Wxx 的sheet")
    week_sheets.sort(key=lambda x: x[0], reverse=True)
    latest_week_num, latest_sheet = week_sheets[0]
    return latest_week_num, latest_sheet["sheet_id"]


def col_to_letter(n):
    """把列数转成Excel列字母，比如 23 -> W"""
    letters = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def get_sheet_data(access_token, spreadsheet_token, sheet_id):
    """读取指定sheet的全部数据，返回DataFrame"""
    headers = {"Authorization": f"Bearer {access_token}"}

    # 先查这个sheet的行列范围
    meta_url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}"
    meta_resp = requests.get(meta_url, headers=headers)
    meta = meta_resp.json()["data"]["sheet"]
    row_count = meta["grid_properties"]["row_count"]
    col_count = meta["grid_properties"]["column_count"]

    last_col = col_to_letter(col_count)
    range_str = f"{sheet_id}!A1:{last_col}{row_count}"
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range_str}"
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"读取数据失败: {data}")

    values = data["data"]["valueRange"]["values"]
    header = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)
    return df


# ========== 数据处理与分析 ==========

def clean_and_analyze(df):
    """清洗原始数据并计算GL大区的各项指标"""
    # 删除无关列
    drop_cols = ['日期', '被抢', '站点断更', '站点断更率', '备注']
    df = df.drop(columns=[c for c in drop_cols if c in df.columns])

    # 数值列从飞书读出来可能是字符串，需要先转数值类型
    numeric_cols = ['应领件量', '领件量', '票数', '1stDay妥投量', '虚假签收量', '妥投量', 'DSP断更']
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce')

    # 填充合并单元格
    fill_cols = ['大区', '中心', 'DSP', '对象']
    df[fill_cols] = df[fill_cols].ffill()

    df_gl = (
        df[df['大区'] == 'GL']
        .groupby(['中心', '对象', 'DSP'])
        .agg(
            总应领件量=('应领件量', 'sum'),
            总领件量=('领件量', 'sum'),
            总票数=('票数', 'sum'),
            总1stDay妥投量=('1stDay妥投量', 'sum'),
            总虚假签收=('虚假签收量', 'sum'),
            总妥投量=('妥投量', 'sum'),
            总DSP断更=('DSP断更', 'sum'),
        )
        .reset_index()
    )

    df_gl['领件率'] = df_gl['总领件量'] / df_gl['总应领件量']
    df_gl['1stDay妥投率'] = df_gl['总1stDay妥投量'] / df_gl['总票数']
    df_gl['DNR率'] = df_gl['总虚假签收'] / df_gl['总妥投量']
    df_gl['DSP断更率'] = df_gl['总DSP断更'] / df_gl['总妥投量']

    rate_cols = ['领件率', '1stDay妥投率', 'DNR率', 'DSP断更率']
    df_gl[rate_cols] = (df_gl[rate_cols] * 100).round(2)
    for col in rate_cols:
        df_gl[col] = df_gl[col].astype(str) + '%'

    keep_cols = (
        ['中心', '对象', 'DSP', '总票数', '总1stDay妥投量', '总妥投量', '总虚假签收', '总DSP断更']
        + rate_cols
    )
    return df_gl[keep_cols]


# ========== 主流程 ==========

def main():
    access_token = get_tenant_access_token()
    spreadsheet_token = get_spreadsheet_token(access_token, WIKI_NODE_TOKEN)
    print("电子表格token:", spreadsheet_token)

    sheets = get_sheet_list(access_token, spreadsheet_token)
    latest_week_num, sheet_id = get_latest_week_sheet(sheets)
    print(f"最新一周是 W{latest_week_num}, sheet_id = {sheet_id}")

    df_raw = get_sheet_data(access_token, spreadsheet_token, sheet_id)
    print(df_raw.head())

    df_gl = clean_and_analyze(df_raw)

    output_file = f"GL路区分析_W{latest_week_num}.xlsx"
    df_gl.to_excel(output_file, index=False)
    print(f"分析完成，已导出: {output_file}")


if __name__ == "__main__":
    main()
