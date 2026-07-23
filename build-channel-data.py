import json
from collections import Counter
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "source-data" / "分贝通明细全年_2026-07-23_1618.xlsx"
MONTHS = [f"{i}月" for i in range(1, 7)]
TRAVEL_TYPES = {"国内机票", "火车", "酒店"}
MGMT_SUPPORT = {"董乾", "杨巍巍", "熊楠星", "李长玉", "岳家璇"}
BIZ_DEV = {"王奉禄", "李铖杰"}


def num(value):
    if pd.isna(value):
        return 0.0
    return float(value)


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def area_from_department(department):
    if "北部" in department:
        return "北部大区"
    if "南部" in department:
        return "南部大区"
    if "东部" in department:
        return "东部大区"
    if "西部" in department:
        return "西部大区"
    return "部门负责人"


def top_items(counter, limit=6):
    return [{"name": name, "value": round(value, 2)} for name, value in counter.most_common(limit)]


detail = pd.read_excel(SOURCE, sheet_name="渠道业务部部分贝通明细")
travel_quota = pd.read_excel(SOURCE, sheet_name="差旅额度")
car_quota = pd.read_excel(SOURCE, sheet_name="用车额度")
collection = pd.read_excel(SOURCE, sheet_name="渠道城市经理月度回款")

detail = detail[detail["月份"].isin(MONTHS)].copy()
detail["企业支付总金额"] = pd.to_numeric(detail["企业支付总金额"], errors="coerce").fillna(0)
detail["人员"] = detail["预订|下单|用餐人"].map(clean)
detail["部门"] = detail["预订|下单人层级部门"].map(clean)
detail["业务线"] = detail["业务线"].map(clean)

travel_quota = travel_quota[travel_quota["部门"].astype(str).str.contains("渠道业务部", na=False)].copy()
car_quota = car_quota[car_quota["部门"].astype(str).str.contains("渠道业务部", na=False)].copy()
travel_by_name = {clean(r["姓名"]): r for _, r in travel_quota.iterrows()}
car_by_name = {clean(r["姓名"]): r for _, r in car_quota.iterrows()}

collection_people = {clean(x) for x in collection["城市经理"].dropna()}
collection_by_name = {clean(r["城市经理"]): r for _, r in collection.iterrows()}
expense_people = {x for x in detail["人员"].unique() if x}
all_people = sorted(expense_people | collection_people)

people = []
for name in all_people:
    rows = detail[detail["人员"] == name]
    department = clean(rows["部门"].mode().iloc[0]) if not rows.empty and not rows["部门"].mode().empty else ""
    tq = travel_by_name.get(name)
    cq = car_by_name.get(name)
    quota_department = clean(tq["部门"]) if tq is not None else clean(cq["部门"]) if cq is not None else ""
    area = clean(collection_by_name[name]["区域"]) if name in collection_by_name else area_from_department(quota_department or department)

    if name in collection_people:
        role = "直接回款"
    elif name in MGMT_SUPPORT:
        role = "区总/管理支持"
    elif name in BIZ_DEV:
        role = "招商拓展"
    else:
        role = "其他/历史人员"

    q2_travel_quota = num(tq.get("二季度额度（5-7月）", 0)) if tq is not None else 0
    q2_car_quota = num(cq.get("2季度用车（5-7月）", 0)) if cq is not None else 0
    monthly = []
    total_travel = 0.0
    total_car = 0.0
    q2_travel_actual = 0.0
    q2_car_actual = 0.0
    for idx, month in enumerate(MONTHS, 1):
        mr = rows[rows["月份"] == month]
        travel_actual = float(mr[mr["业务线"].isin(TRAVEL_TYPES)]["企业支付总金额"].sum())
        car_actual = float(mr[mr["业务线"] == "用车"]["企业支付总金额"].sum())
        total_actual = travel_actual + car_actual
        total_travel += travel_actual
        total_car += car_actual
        if idx >= 5:
            q2_travel_actual += travel_actual
            q2_car_actual += car_actual
        travel_month_quota = num(tq.get(f"{idx}月差旅", 0)) if tq is not None and idx <= 4 else q2_travel_quota / 3
        car_month_quota = num(cq.get(f"{idx}月用车", 0)) if cq is not None and idx <= 4 else q2_car_quota / 3
        monthly.append(
            {
                "month": month,
                "travelActual": round(travel_actual, 2),
                "carActual": round(car_actual, 2),
                "totalActual": round(total_actual, 2),
                "travelQuota": round(travel_month_quota, 2),
                "carQuota": round(car_month_quota, 2),
                "travelUsage": travel_actual / travel_month_quota if travel_month_quota else None,
                "carUsage": car_actual / car_month_quota if car_month_quota else None,
            }
        )

    jan_apr_travel_quota = sum(num(tq.get(f"{i}月差旅", 0)) for i in range(1, 5)) if tq is not None else 0
    jan_apr_car_quota = sum(num(cq.get(f"{i}月用车", 0)) for i in range(1, 5)) if cq is not None else 0
    period_travel_quota = jan_apr_travel_quota + q2_travel_quota / 3 * 2
    period_car_quota = jan_apr_car_quota + q2_car_quota / 3 * 2

    collection_monthly = []
    collection_total = 0.0
    target_total = 0.0
    cr = collection_by_name.get(name)
    for idx, month in enumerate(MONTHS, 1):
        if cr is None:
            col = None
            target = None
        else:
            col = num(cr.get(f"{idx}月回款", 0))
            target = num(cr.get(f"{idx}月目标", 0))
            collection_total += col
            target_total += target
        collection_monthly.append(
            {
                "month": month,
                "collection": round(col, 2) if col is not None else None,
                "target": round(target, 2) if target is not None else None,
                "achievement": col / target if col is not None and target else None,
            }
        )

    total_cost = total_travel + total_car
    q2_total_quota = q2_travel_quota + q2_car_quota
    q2_total_actual = q2_travel_actual + q2_car_actual
    period_total_quota = period_travel_quota + period_car_quota
    period_usage = total_cost / period_total_quota if period_total_quota else None
    roi = collection_total / total_cost if total_cost else None
    achievement = collection_total / target_total if target_total else None
    expense_rate = total_cost / collection_total if collection_total else None

    city_counter = Counter()
    route_counter = Counter()
    route_count_counter = Counter()
    type_counter = Counter()
    for _, row in rows.iterrows():
        amount = num(row["企业支付总金额"])
        destination = clean(row["到达城市|入住城市|目的城市|用餐城市|还车城市"])
        origin = clean(row["出发城市|下单城市|取车城市"])
        if destination:
            city_counter[destination] += abs(amount)
        if origin and destination:
            route_counter[f"{origin}-{destination}"] += abs(amount)
            route_count_counter[f"{origin}-{destination}"] += 1
        type_counter[clean(row["业务线"])] += amount

    negative_rows = rows[rows["企业支付总金额"] < 0]
    record_count = len(rows)
    negative_record_count = len(negative_rows)
    negative_amount = abs(float(negative_rows["企业支付总金额"].sum()))

    score = 0
    tags = []
    if role == "直接回款":
        if achievement is not None and achievement < 0.7:
            score += 5
            tags.append("回款低于70%")
        if expense_rate is not None and expense_rate > 0.08:
            score += 3
            tags.append("费用率偏高")
        if roi is not None and roi < 10:
            score += 3
            tags.append("投入产出偏低")
    q2_usage = q2_total_actual / q2_total_quota if q2_total_quota else None
    if period_usage is not None and period_usage > 1:
        score += 2
        tags.append("1–6月额度超额")
    elif period_usage is not None and period_usage > 0.8:
        score += 2
        tags.append("1–6月额度预警")
    if len(city_counter) >= 4:
        tags.append("到访城市多")

    if score >= 7:
        risk = "高"
    elif score >= 3:
        risk = "中"
    else:
        risk = "低"
    if role != "直接回款":
        category = "非直接回款角色"
    elif achievement is not None and achievement >= 1 and (roi or 0) >= 10:
        category = "高效型"
    elif achievement is not None and achievement < 0.7:
        category = "高费用低回款"
    else:
        category = "稳健型"

    people.append(
        {
            "name": name,
            "department": department or quota_department,
            "area": area,
            "roleType": role,
            "category": category,
            "risk": risk,
            "riskScore": score,
            "travelActual": round(total_travel, 2),
            "carActual": round(total_car, 2),
            "totalActual": round(total_cost, 2),
            "janAprTravelQuota": round(jan_apr_travel_quota, 2),
            "janAprCarQuota": round(jan_apr_car_quota, 2),
            "periodTravelQuota": round(period_travel_quota, 2),
            "periodCarQuota": round(period_car_quota, 2),
            "periodTotalQuota": round(period_total_quota, 2),
            "periodTravelUsage": total_travel / period_travel_quota if period_travel_quota else None,
            "periodCarUsage": total_car / period_car_quota if period_car_quota else None,
            "periodUsage": period_usage,
            "periodTravelRemaining": round(period_travel_quota - total_travel, 2),
            "periodCarRemaining": round(period_car_quota - total_car, 2),
            "periodRemaining": round(period_total_quota - total_cost, 2),
            "q2TravelQuota": round(q2_travel_quota, 2),
            "q2CarQuota": round(q2_car_quota, 2),
            "q2TravelActual": round(q2_travel_actual, 2),
            "q2CarActual": round(q2_car_actual, 2),
            "q2TotalQuota": round(q2_total_quota, 2),
            "q2TotalActual": round(q2_total_actual, 2),
            "q2TravelUsage": q2_travel_actual / q2_travel_quota if q2_travel_quota else None,
            "q2CarUsage": q2_car_actual / q2_car_quota if q2_car_quota else None,
            "q2Usage": q2_usage,
            "q2Remaining": round(q2_total_quota - q2_total_actual, 2),
            "collection": round(collection_total, 2),
            "target": round(target_total, 2),
            "achievement": achievement,
            "roi": roi,
            "expenseRate": expense_rate,
            "monthly": monthly,
            "collectionMonthly": collection_monthly,
            "topCities": top_items(city_counter),
            "topRoutes": top_items(route_counter, 4),
            "topRoutesByCount": [
                {"name": route, "value": count}
                for route, count in route_count_counter.most_common(4)
            ],
            "recordCount": record_count,
            "averagePerRecord": total_cost / record_count if record_count else None,
            "negativeRecordCount": negative_record_count,
            "negativeAmount": round(negative_amount, 2),
            "cityCount": len(city_counter),
            "routeCount": len(route_count_counter),
            "byType": {k: round(v, 2) for k, v in type_counter.items() if k},
            "tags": tags or ["暂无异常"],
        }
    )

monthly_summary = []
for idx, month in enumerate(MONTHS, 1):
    rows = detail[detail["月份"] == month]
    travel_actual = float(rows[rows["业务线"].isin(TRAVEL_TYPES)]["企业支付总金额"].sum())
    car_actual = float(rows[rows["业务线"] == "用车"]["企业支付总金额"].sum())
    col = sum(
        p["collectionMonthly"][idx - 1]["collection"] or 0
        for p in people
        if p["roleType"] == "直接回款"
    )
    target = sum(
        p["collectionMonthly"][idx - 1]["target"] or 0
        for p in people
        if p["roleType"] == "直接回款"
    )
    monthly_summary.append(
        {
            "month": month,
            "travelActual": round(travel_actual, 2),
            "carActual": round(car_actual, 2),
            "totalActual": round(travel_actual + car_actual, 2),
            "travelQuota": round(sum(p["monthly"][idx - 1]["travelQuota"] for p in people), 2),
            "carQuota": round(sum(p["monthly"][idx - 1]["carQuota"] for p in people), 2),
            "collection": round(col, 2),
            "target": round(target, 2),
            "achievement": col / target if target else None,
            "roi": col / (travel_actual + car_actual) if travel_actual + car_actual else None,
        }
    )

area_summary = []
for area in sorted({p["area"] for p in people}):
    members = [p for p in people if p["area"] == area]
    area_summary.append(
        {
            "area": area,
            "people": len(members),
            "travelActual": round(sum(p["travelActual"] for p in members), 2),
            "carActual": round(sum(p["carActual"] for p in members), 2),
            "totalActual": round(sum(p["totalActual"] for p in members), 2),
            "periodTravelQuota": round(sum(p["periodTravelQuota"] for p in members), 2),
            "periodCarQuota": round(sum(p["periodCarQuota"] for p in members), 2),
            "periodTotalQuota": round(sum(p["periodTotalQuota"] for p in members), 2),
            "periodTravelRemaining": round(sum(p["periodTravelRemaining"] for p in members), 2),
            "periodCarRemaining": round(sum(p["periodCarRemaining"] for p in members), 2),
            "periodRemaining": round(sum(p["periodRemaining"] for p in members), 2),
            "q2TravelQuota": round(sum(p["q2TravelQuota"] for p in members), 2),
            "q2CarQuota": round(sum(p["q2CarQuota"] for p in members), 2),
            "q2TotalQuota": round(sum(p["q2TotalQuota"] for p in members), 2),
            "q2Actual": round(sum(p["q2TotalActual"] for p in members), 2),
            "q2Remaining": round(sum(p["q2Remaining"] for p in members), 2),
        }
    )
    area_summary[-1]["q2Usage"] = (
        area_summary[-1]["q2Actual"] / area_summary[-1]["q2TotalQuota"]
        if area_summary[-1]["q2TotalQuota"]
        else None
    )
    area_summary[-1]["periodTravelUsage"] = (
        area_summary[-1]["travelActual"] / area_summary[-1]["periodTravelQuota"]
        if area_summary[-1]["periodTravelQuota"]
        else None
    )
    area_summary[-1]["periodCarUsage"] = (
        area_summary[-1]["carActual"] / area_summary[-1]["periodCarQuota"]
        if area_summary[-1]["periodCarQuota"]
        else None
    )
    area_summary[-1]["periodUsage"] = (
        area_summary[-1]["totalActual"] / area_summary[-1]["periodTotalQuota"]
        if area_summary[-1]["periodTotalQuota"]
        else None
    )

total_actual = sum(x["totalActual"] for x in monthly_summary)
total_travel = sum(x["travelActual"] for x in monthly_summary)
total_car = sum(x["carActual"] for x in monthly_summary)
total_collection = sum((x["collection"] or 0) for x in monthly_summary)
total_target = sum((x["target"] or 0) for x in monthly_summary)

result = {
    "generatedAt": "2026-07-23",
    "source": "飞书多维表格：分贝通明细全年",
    "period": "2026年1-6月",
    "notes": {
        "expense": "企业支付净额，全状态正负冲抵",
        "collection": "城市经理回款与目标数据覆盖1–6月",
        "quota": "1–4月使用月度额度；5–6月按5–7月季度额度÷3折算月均额度",
    },
    "summary": {
        "people": len(people),
        "expensePeople": len(expense_people),
        "directPeople": len(collection_people),
        "totalActual": round(total_actual, 2),
        "travelActual": round(total_travel, 2),
        "carActual": round(total_car, 2),
        "collection": round(total_collection, 2),
        "target": round(total_target, 2),
        "achievement": total_collection / total_target if total_target else None,
        "roi": total_collection / total_actual if total_actual else None,
        "expenseRate": total_actual / total_collection if total_collection else None,
        "highRisk": sum(1 for p in people if p["risk"] == "高"),
        "efficient": sum(1 for p in people if p["category"] == "高效型"),
        "periodTravelQuota": round(sum(p["periodTravelQuota"] for p in people), 2),
        "periodCarQuota": round(sum(p["periodCarQuota"] for p in people), 2),
        "periodTotalQuota": round(sum(p["periodTotalQuota"] for p in people), 2),
        "periodTravelRemaining": round(sum(p["periodTravelRemaining"] for p in people), 2),
        "periodCarRemaining": round(sum(p["periodCarRemaining"] for p in people), 2),
        "periodRemaining": round(sum(p["periodRemaining"] for p in people), 2),
        "q2TravelQuota": round(sum(p["q2TravelQuota"] for p in people), 2),
        "q2CarQuota": round(sum(p["q2CarQuota"] for p in people), 2),
        "q2TotalQuota": round(sum(p["q2TotalQuota"] for p in people), 2),
        "q2Actual": round(sum(p["q2TotalActual"] for p in people), 2),
        "q2Remaining": round(sum(p["q2Remaining"] for p in people), 2),
    },
    "monthlySummary": monthly_summary,
    "areaSummary": area_summary,
    "people": people,
}
result["summary"]["q2Usage"] = (
    result["summary"]["q2Actual"] / result["summary"]["q2TotalQuota"]
    if result["summary"]["q2TotalQuota"]
    else None
)
result["summary"]["periodTravelUsage"] = (
    result["summary"]["travelActual"] / result["summary"]["periodTravelQuota"]
    if result["summary"]["periodTravelQuota"]
    else None
)
result["summary"]["periodCarUsage"] = (
    result["summary"]["carActual"] / result["summary"]["periodCarQuota"]
    if result["summary"]["periodCarQuota"]
    else None
)
result["summary"]["periodUsage"] = (
    result["summary"]["totalActual"] / result["summary"]["periodTotalQuota"]
    if result["summary"]["periodTotalQuota"]
    else None
)

(ROOT / "data.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(
    json.dumps(
        {
            "people": result["summary"]["people"],
            "totalActual": result["summary"]["totalActual"],
            "travelActual": result["summary"]["travelActual"],
            "carActual": result["summary"]["carActual"],
            "collection": result["summary"]["collection"],
            "target": result["summary"]["target"],
            "q2Quota": result["summary"]["q2TotalQuota"],
            "q2Actual": result["summary"]["q2Actual"],
        },
        ensure_ascii=False,
    )
)
