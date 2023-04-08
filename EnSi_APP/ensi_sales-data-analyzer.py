import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from datetime import datetime, timedelta
import pandas as pd

# 选择Excel文件
def select_excel_file(title):
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])
    return file_path

# 读取销售数据报表
sales_data_path = select_excel_file("选择销售数据")
sales_data = pd.read_excel(sales_data_path, parse_dates=["日期"])

# 计算昨天的日期（数据源中的最后一天）
yesterday = sales_data["日期"].max().strftime("%Y-%m-%d")

# 读取业务员销售目标表
sales_target_path = select_excel_file("选择销售目标")
sales_target = pd.read_excel(sales_target_path)

# 生成完成率excel报表
def generate_target_completion_report(sales_data, sales_target, target_type):
    sales_this_month = sales_data[sales_data["日期"].dt.month == pd.to_datetime(yesterday).month]

    # Calculate the sum of sales for each employee this month
    sales_sum = sales_this_month.groupby("店员")["实销数"].sum()

    # Merge the sales sum with the sales target DataFrame
    target_completion = sales_target.merge(sales_sum, how='left', left_on='店员', right_index=True).fillna(0)
    
    target_completion["已完成"] = target_completion["实销数"].round().astype(int)
    target_completion["待完成"] = (target_completion[target_type] - target_completion["已完成"]).round().astype(int)
    target_completion.loc[target_completion["待完成"] < 0, "待完成"] = "已完成"

    target_completion["完成率"] = target_completion["已完成"] / target_completion[target_type]

    # Convert completion rate to a float for sorting purposes
    target_completion["完成率_float"] = target_completion["完成率"]
    
    # Round the target_type values to whole numbers
    target_completion[target_type] = target_completion[target_type].round().astype(int)

    target_completion = target_completion[["店员", target_type, "已完成", "待完成", "完成率", "完成率_float"]].reset_index(drop=True)

    # Sort the DataFrame by completion rate
    target_completion.sort_values(by="完成率_float", ascending=False, inplace=True)

    # Reset the index to represent the completion rate ranking
    target_completion.reset_index(drop=True, inplace=True)
    target_completion.index += 1
    target_completion.index.name = "序号"

    # Format the completion rate column as a percentage string
    target_completion["完成率"] = target_completion["完成率"].apply(lambda x: f"{x*100:.2f}%" if x <= 1 else f"{x*100:.2f}%")

    # Remove the temporary float column
    target_completion.drop(columns=["完成率_float"], inplace=True)

    return target_completion

def save_target_completion_report():
    target_type = target_level.get()
    target_completion = generate_target_completion_report(sales_data, sales_target, target_type)

    # Dynamically set the output file name based on the user's selection
    today = datetime.now().strftime("%Y-%m-%d")
    output_file = f"各档口个人销量{target_type}完成度（{today}）.xlsx"
    
    target_completion.to_excel(output_file, index=True)

# GUI for selecting target level
def select_target_level():
    global target_level
    target_level_window = tk.Toplevel(root)
    target_level_window.title("选择目标级别")

    ttk.Label(target_level_window, text="请选择目标级别:").grid(column=0, row=0)

    target_level = tk.StringVar()
    target_level.set("月度低标")

    ttk.OptionMenu(target_level_window, target_level, *["月度低标", "月度中标", "月度高标"]).grid(column=1, row=0)

    ttk.Button(target_level_window, text="生成报表", command=save_target_completion_report).grid(column=1, row=1)
    
root = tk.Tk()
root.withdraw()

# 数据处理和排行榜生成
def generate_rankings(sales_data, sales_target, yesterday):
    # 计算昨日和本月的数据
    sales_yesterday = sales_data[sales_data["日期"] == yesterday]
    sales_this_month = sales_data[sales_data["日期"].dt.month == pd.to_datetime(yesterday).month]

    # 昨日个人销量排名
    ranking1 = sales_yesterday.groupby("店员")["实销数"].sum().sort_values(ascending=False)
    ranking1_numeric = ranking1.copy()
    
    for clerk in ranking1.index:
        clerk_target = sales_target[sales_target["店员"] == clerk]
        if not clerk_target.empty:
            if ranking1.loc[clerk] >= clerk_target["日高标"].values[0]: # Check for '日高标' first
                ranking1.loc[clerk] = f"{ranking1.loc[clerk]} 完成高标"
            elif ranking1.loc[clerk] >= clerk_target["日中标"].values[0]:
                ranking1.loc[clerk] = f"{ranking1.loc[clerk]} 完成中标"
            elif ranking1.loc[clerk] >= clerk_target["日低标"].values[0]:
                ranking1.loc[clerk] = f"{ranking1.loc[clerk]} 完成低标"

    # 本月个人总销量排名
    ranking2 = sales_this_month.groupby("店员")["实销数"].sum().sort_values(ascending=False)

    # 昨日个人实收排名
    ranking3 = sales_yesterday.groupby("店员")["实收"].sum().sort_values(ascending=False)

    # 当月累计实收排名
    ranking4 = sales_this_month.groupby("店员")["实收"].sum().sort_values(ascending=False)

    # 本月个人业绩低标达成率排名
    sales_target_index = sales_target.set_index("店员")
    sales_sum = sales_this_month.groupby("店员")["实销数"].sum()
    target_low = sales_sum.div(sales_target_index["月度低标"]).dropna()
    ranking5 = target_low.sort_values(ascending=False)

    # 本月个人业绩高标达成率排名
    target_high = sales_sum.div(sales_target_index["月度高标"]).dropna()
    ranking6 = target_high.sort_values(ascending=False)

    # 昨日档口销量排名
    ranking7 = sales_yesterday.groupby("门店")["实销数"].sum().sort_values(ascending=False)

    return ranking1, ranking2, ranking3, ranking4, ranking5, ranking6, ranking7, ranking1_numeric

rankings = generate_rankings(sales_data, sales_target, yesterday)

# 计算团队目标完成率
def generate_team_performance(sales_data, sales_target, yesterday):
    team_performance = []

    # 计算昨日的数据
    sales_yesterday = sales_data[sales_data["日期"] == yesterday]
    
    # Group by store and calculate the sum of targets
    store_targets = sales_target.groupby("门店").sum(numeric_only=True)
    store_sales = sales_data.groupby(["门店", "店员"])["实销数"].sum().reset_index().groupby("门店").sum(numeric_only=True)

    for store in store_targets.index:
        store_target_high = store_targets.loc[store, "月度高标"] 
        store_target_low = store_targets.loc[store, "月度低标"] 
        store_sales_sum = store_sales.loc[store, "实销数"]

        low_completion_rate = store_sales_sum / store_target_low * 100
        high_completion_rate = store_sales_sum / store_target_high * 100

        yesterday_target_low = store_targets.loc[store, "日低标"]
        yesterday_sales_sum = sales_yesterday[sales_yesterday["门店"] == store]["实销数"].sum()
        
        store_performance = (
            f"{store}\n"
            f"月度高标： {round(store_target_high):.1f}件\n"
            f"月度低标： {round(store_target_low):.1f}件\n"
            f"已完成： {round(store_sales_sum):.1f}件\n"
            f"低标完成率 {low_completion_rate:.1f}%\n"
            f"高标完成率 {high_completion_rate:.1f}%\n\n"
            f"昨日团队完成情况\n"
            f"昨日低标： {round(yesterday_target_low):.1f}件\n"
            f"实际完成： {round(yesterday_sales_sum):.1f}件\n"
            f"{'-' * 35}\n"
        )

        team_performance.append(store_performance)

    return team_performance

# 导出排行榜到TXT文件
def export_rankings_to_txt(rankings, team_performance):
    today = datetime.now().strftime("%Y-%m-%d")
    output_file = f"每日销售排行榜（{today}）.txt"
    separator = "-" * 50
    ranking_titles = [
        "昨日个人销量排名",
        "本月个人总销量排名",
        "昨日个人实收排名",
        "当月累计实收排名",
        "本月个人业绩低标达成率排名",
        "本月个人业绩高标达成率排名",
        "昨日档口销量排名",
    ]

    with open(output_file, "w", encoding="utf-8") as f:
        for index, (ranking, title) in enumerate(zip(rankings[:-1], ranking_titles), start=1): # Exclude ranking3_numeric from the loop
            # Calculate the sum only for numeric values
            numeric_sum = ranking[ranking.apply(lambda x: isinstance(x, (int, float)))].sum()

            if index == 1: # Modify the total sum calculation for 昨日个人销量排名
                numeric_sum = rankings[-1].sum() # Use ranking3_numeric for the total sum calculation
                
            if index == 4:
                numeric_sum /= 10000  # Convert sum to 万 for 当月累计实收排名
                ranking = ranking / 10000  # Divide the ranking values by 10000 to get the value in 万
                ranking = ranking.round(1)  # Round the values to 1 decimal place
                ranking = ranking.apply(lambda x: f"{x}万")  # Add '万' after each value
                sum_suffix = "万"
            elif index in [5, 6]:
                sum_suffix = "%"
                numeric_sum = (numeric_sum / len(ranking)) * 100  # Calculate average percentage
            elif index in [1, 2, 7]:
                sum_suffix = "件"
            else:
                sum_suffix = "元"
                
            # Add ranking column
            ranking_with_rank = ranking.reset_index()
            ranking_with_rank.index += 1
            ranking_with_rank.index.name = None

            # Convert low and high target achievement rates to percentages
            if index in [5, 6]:
                if ranking.name:
                    ranking_with_rank[ranking.name] = ranking_with_rank[ranking.name].apply(lambda x: f"{x*100:.2f}%" if isinstance(x, (int, float)) else x)
                else:
                    ranking_with_rank.iloc[:, 1] = ranking_with_rank.iloc[:, 1].apply(lambda x: f"{x*100:.2f}%" if isinstance(x, (int, float)) else x)

            f.write(f"{title}\n\n")
            f.write(ranking_with_rank.to_string(header=False, index=True))

            f.write(f"\n\n合计：{numeric_sum:.2f}{sum_suffix}\n\n")
            f.write(separator)
            f.write("\n")
        f.write("\n团队目标达成情况\n\n")
        for store_performance in team_performance:
            f.write(store_performance)

team_performance = generate_team_performance(sales_data, sales_target, yesterday)
export_rankings_to_txt(rankings, team_performance)

root.deiconify()
root.title("报表生成工具")

ttk.Button(root, text="选择目标级别", command=select_target_level).grid(column=0, row=1)

root.mainloop()
