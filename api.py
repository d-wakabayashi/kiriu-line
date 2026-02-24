"""
KIRIU ライン負荷最適化システム - Web API
Google Apps Scriptから呼び出すためのFastAPI

デプロイ先: Cloud Run / Render / Railway など
"""

import json
from typing import Any
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from config import (
    DISC_LINES, DEFAULT_CAPACITIES, MONTHS,
    DEFAULT_JPH, DEFAULT_WORK_PATTERNS, DEFAULT_MONTHLY_WORKING_DAYS,
    WorkPattern, calculate_monthly_capacities,
)
from model import optimize, OptimizationResult
from data_loader import PartSpec, PartDemand

app = FastAPI(
    title="KIRIU ライン負荷最適化API",
    description="生産ラインの負荷を最適化するAPI",
    version="1.0.0",
)

# CORS設定（GASからのアクセスを許可）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class PartInput(BaseModel):
    """部品入力データ"""
    part_number: str
    part_name: str = ""
    main_line: str
    sub1_line: str | None = None
    sub2_line: str | None = None
    monthly_demand: list[int]  # 12ヶ月分


class OptimizeRequest(BaseModel):
    """最適化リクエスト"""
    parts: list[PartInput]
    capacities: dict[str, int | list[int]] | None = None  # 月別能力対応
    time_limit: int = 60
    load_rate_limit: float = 1.0  # 負荷率上限（0.0〜1.0）


class LineLoadOutput(BaseModel):
    """ライン負荷出力"""
    line: str
    monthly_capacities: list[int]  # 月別能力
    monthly_loads: list[int]
    avg_capacity: float
    avg_load: float
    load_rate: float


class UnmetDemandOutput(BaseModel):
    """未割当出力"""
    part_number: str
    monthly_unmet: list[int]  # 月別未割当
    total_unmet: int


class PartAllocationOutput(BaseModel):
    """部品割当出力"""
    part_number: str
    allocations: dict[str, list[int]]  # {ライン: [月別数量]}


class OptimizeResponse(BaseModel):
    """最適化レスポンス"""
    success: bool
    status: str
    objective_value: float | None
    solve_time: float
    line_loads: list[LineLoadOutput]
    allocations: list[PartAllocationOutput]
    unmet_demands: list[UnmetDemandOutput]  # 未割当一覧
    summary: dict[str, Any]


@app.get("/")
def root():
    """ヘルスチェック"""
    return {"status": "ok", "service": "KIRIU Line Optimizer API"}


@app.get("/lines")
def get_lines():
    """利用可能なライン一覧を取得"""
    return {
        "lines": DISC_LINES,
        "default_capacities": DEFAULT_CAPACITIES,
        "months": MONTHS,
    }


@app.post("/optimize", response_model=OptimizeResponse)
def run_optimization(request: OptimizeRequest):
    """
    最適化を実行

    リクエスト例:
    ```json
    {
        "parts": [
            {
                "part_number": "PART001",
                "part_name": "部品A",
                "main_line": "4915",
                "sub1_line": "4919",
                "monthly_demand": [1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000]
            }
        ],
        "capacities": {"4915": 70000, "4919": 80000, ...},
        "time_limit": 60
    }
    ```
    """
    # 入力データを変換
    specs = {}
    demands = {}

    for part in request.parts:
        if part.main_line not in DISC_LINES:
            raise HTTPException(
                status_code=400,
                detail=f"無効なメインライン: {part.main_line}。有効なライン: {DISC_LINES}"
            )

        if len(part.monthly_demand) != 12:
            raise HTTPException(
                status_code=400,
                detail=f"部品 {part.part_number} の月別需要は12ヶ月分必要です"
            )

        specs[part.part_number] = PartSpec(
            part_number=part.part_number,
            part_name=part.part_name,
            main_line=part.main_line,
            sub1_line=part.sub1_line if part.sub1_line in DISC_LINES else None,
            sub2_line=part.sub2_line if part.sub2_line in DISC_LINES else None,
        )

        demands[part.part_number] = PartDemand(
            part_number=part.part_number,
            part_name=part.part_name,
            monthly_demand=part.monthly_demand,
        )

    # 能力設定
    capacities = request.capacities or DEFAULT_CAPACITIES.copy()

    # 不足しているラインのデフォルト能力を追加
    for line in DISC_LINES:
        if line not in capacities:
            capacities[line] = DEFAULT_CAPACITIES.get(line, 50000)

    # 最適化実行
    try:
        result = optimize(specs, demands, capacities, request.time_limit, request.load_rate_limit)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"最適化エラー: {str(e)}")

    # 月別能力を正規化
    def normalize_caps(caps):
        result = {}
        for line in DISC_LINES:
            cap = caps.get(line, DEFAULT_CAPACITIES.get(line, 0))
            if isinstance(cap, list):
                result[line] = (cap + [cap[-1]] * 12)[:12] if cap else [0] * 12
            else:
                result[line] = [cap] * 12
        return result

    monthly_capacities = normalize_caps(capacities)

    # 結果を整形
    line_loads = []
    for line in DISC_LINES:
        loads = result.line_loads.get(line, [0] * 12)
        line_caps = monthly_capacities.get(line, [0] * 12)
        avg_cap = sum(line_caps) / 12
        avg_load = sum(loads) / 12
        load_rate = avg_load / avg_cap if avg_cap > 0 else 0

        line_loads.append(LineLoadOutput(
            line=line,
            monthly_capacities=line_caps,
            monthly_loads=loads,
            avg_capacity=avg_cap,
            avg_load=avg_load,
            load_rate=load_rate,
        ))

    allocations = []
    for part_num, alloc in result.allocation.items():
        allocations.append(PartAllocationOutput(
            part_number=part_num,
            allocations=alloc,
        ))

    # 未割当一覧
    unmet_demands = []
    if result.unmet_demand:
        for part_num, monthly_unmet in result.unmet_demand.items():
            if sum(monthly_unmet) > 0:
                unmet_demands.append(UnmetDemandOutput(
                    part_number=part_num,
                    monthly_unmet=monthly_unmet,
                    total_unmet=sum(monthly_unmet),
                ))

    # サマリー
    total_demand = sum(sum(d.monthly_demand) for d in demands.values())
    total_capacity = sum(sum(caps) for caps in monthly_capacities.values())
    total_unmet = sum(u.total_unmet for u in unmet_demands)

    summary = {
        "total_parts": len(specs),
        "total_demand": total_demand,
        "total_capacity": total_capacity,
        "overall_load_rate": total_demand / total_capacity if total_capacity > 0 else 0,
        "total_unmet": total_unmet,
        "unmet_parts_count": len(unmet_demands),
    }

    return OptimizeResponse(
        success=result.status in ('OPTIMAL', 'FEASIBLE'),
        status=result.status,
        objective_value=result.objective_value,
        solve_time=result.solve_time,
        line_loads=line_loads,
        allocations=allocations,
        unmet_demands=unmet_demands,
        summary=summary,
    )


# ============================================================
# シンプル版API（スプレッドシートから直接呼び出し用）
# ============================================================

class SimpleOptimizeRequest(BaseModel):
    """シンプル版リクエスト（2次元配列形式）"""
    # 部品データ: [[部品番号, メインライン, サブ1, サブ2, 4月, 5月, ..., 3月], ...]
    parts_data: list[list[Any]]
    # 能力データ: [[ライン, 4月, 5月, ..., 3月], ...] または [[ライン, 能力], ...]
    capacities_data: list[list[Any]] | None = None
    time_limit: int = 60
    load_rate_limit: float = 1.0  # 負荷率上限（0.0〜1.0）


def _parse_simple_request(request: SimpleOptimizeRequest):
    """シンプル版リクエストからspecs, demands, capacitiesを抽出する共通ヘルパー

    同一部品番号・同一ラインの行が複数ある場合は需要量を合算する。
    """
    specs = {}
    demands = {}

    for row in request.parts_data:
        if len(row) < 16:
            continue

        part_num = str(row[0]).strip()
        if not part_num or part_num == '部品番号':
            continue

        main_line = str(row[1]).strip() if row[1] else None
        sub1_line = str(row[2]).strip() if row[2] else None
        sub2_line = str(row[3]).strip() if row[3] else None

        if not main_line or main_line not in DISC_LINES:
            continue

        monthly = []
        for i in range(4, 16):
            try:
                val = int(float(row[i])) if row[i] else 0
            except (ValueError, TypeError):
                val = 0
            monthly.append(max(0, val))

        if sum(monthly) == 0:
            continue

        # 仕様を登録（最初に見つかったものを使用）
        if part_num not in specs:
            specs[part_num] = PartSpec(
                part_number=part_num,
                part_name='',
                main_line=main_line,
                sub1_line=sub1_line if sub1_line in DISC_LINES else None,
                sub2_line=sub2_line if sub2_line in DISC_LINES else None,
            )

        # 同一部品番号の複数行は需要を合算
        if part_num in demands:
            existing = demands[part_num]
            for i in range(12):
                existing.monthly_demand[i] += monthly[i]
        else:
            demands[part_num] = PartDemand(
                part_number=part_num,
                part_name='',
                monthly_demand=monthly,
            )

    if not specs:
        raise HTTPException(status_code=400, detail="有効な部品データがありません")

    capacities = {}
    if request.capacities_data:
        for row in request.capacities_data:
            if len(row) < 2:
                continue
            line = str(row[0]).strip()
            if line not in DISC_LINES or line == 'ライン':
                continue

            if len(row) >= 13:
                monthly_caps = []
                for i in range(1, 13):
                    try:
                        val = int(float(row[i])) if row[i] else 0
                    except (ValueError, TypeError):
                        val = DEFAULT_CAPACITIES.get(line, 50000)
                    monthly_caps.append(max(0, val))
                capacities[line] = monthly_caps
            else:
                try:
                    cap = int(float(row[1]))
                    capacities[line] = cap
                except (ValueError, TypeError):
                    pass

    for line in DISC_LINES:
        if line not in capacities:
            capacities[line] = DEFAULT_CAPACITIES.get(line, 50000)

    return specs, demands, capacities


@app.post("/optimize/simple")
def run_simple_optimization(request: SimpleOptimizeRequest):
    """
    シンプル版最適化（スプレッドシートから直接呼び出し用）

    parts_data形式:
    [
        ["部品番号", "メインライン", "サブ1", "サブ2", 4月, 5月, ..., 3月],
        ["PART001", "4915", "4919", "", 1000, 1000, ..., 1000],
        ...
    ]
    """
    specs, demands, capacities = _parse_simple_request(request)

    # 最適化実行
    result = optimize(specs, demands, capacities, request.time_limit, request.load_rate_limit)

    # 月別能力を正規化
    def normalize_caps(caps):
        result = {}
        for line in DISC_LINES:
            cap = caps.get(line, DEFAULT_CAPACITIES.get(line, 0))
            if isinstance(cap, list):
                result[line] = (cap + [cap[-1]] * 12)[:12] if cap else [0] * 12
            else:
                result[line] = [cap] * 12
        return result

    monthly_capacities = normalize_caps(capacities)

    # スプレッドシート用の2次元配列形式で返す
    # ライン負荷結果（月別能力対応）
    line_loads_array = [["ライン"] + MONTHS + ["平均能力", "平均負荷", "負荷率"]]
    for line in DISC_LINES:
        loads = result.line_loads.get(line, [0] * 12)
        line_caps = monthly_capacities.get(line, [0] * 12)
        avg_cap = sum(line_caps) / 12
        avg_load = sum(loads) / 12
        rate = avg_load / avg_cap if avg_cap > 0 else 0
        line_loads_array.append([line] + loads + [int(avg_cap), int(avg_load), f"{rate:.1%}"])

    # キャパシティ行を追加
    cap_row = ["キャパシティ"] + [monthly_capacities[DISC_LINES[0]][m] for m in range(12)] + ["", "", ""]
    # 各ラインのキャパシティ用に別配列
    capacity_array = [["ライン"] + MONTHS]
    for line in DISC_LINES:
        capacity_array.append([line] + monthly_capacities[line])

    # 部品割当結果
    alloc_array = [["部品番号", "割当ライン"] + MONTHS + ["年間計"]]
    for part_num in sorted(result.allocation.keys()):
        for line, monthly in result.allocation[part_num].items():
            if sum(monthly) > 0:
                alloc_array.append([part_num, line] + monthly + [sum(monthly)])

    # 未割当結果
    unmet_array = [["部品番号"] + MONTHS + ["年間計"]]
    total_unmet = 0
    if result.unmet_demand:
        for part_num in sorted(result.unmet_demand.keys()):
            monthly_unmet = result.unmet_demand[part_num]
            if sum(monthly_unmet) > 0:
                unmet_array.append([part_num] + monthly_unmet + [sum(monthly_unmet)])
                total_unmet += sum(monthly_unmet)

    return {
        "success": result.status in ('OPTIMAL', 'FEASIBLE'),
        "status": result.status,
        "objective_value": result.objective_value,
        "solve_time": result.solve_time,
        "line_loads": line_loads_array,
        "capacities": capacity_array,
        "allocations": alloc_array,
        "unmet_demands": unmet_array,
        "parts_count": len(specs),
        "total_demand": sum(sum(d.monthly_demand) for d in demands.values()),
        "total_unmet": total_unmet,
    }


# ============================================================
# 複数負荷率パターン比較API
# ============================================================

LOAD_RATE_PATTERNS = [1.0, 0.9, 0.8]


class CompareOptimizeRequest(BaseModel):
    """複数パターン比較リクエスト（2次元配列形式）"""
    parts_data: list[list[Any]]
    capacities_data: list[list[Any]] | None = None
    time_limit: int = 60
    load_rate_patterns: list[float] | None = None  # カスタムパターン（省略時は100/90/80%）


@app.post("/optimize/simple/compare")
def run_compare_optimization(request: CompareOptimizeRequest):
    """
    複数負荷率パターンで最適化を実行し、比較結果を返す

    3パターン（100%/90%/80%）で最適化を実行し、
    各パターンの結果と比較サマリーをスプレッドシート用の2次元配列で返す
    """
    # SimpleOptimizeRequestと同じパース処理を利用
    simple_req = SimpleOptimizeRequest(
        parts_data=request.parts_data,
        capacities_data=request.capacities_data,
        time_limit=request.time_limit,
    )
    specs, demands, capacities = _parse_simple_request(simple_req)

    patterns = request.load_rate_patterns or LOAD_RATE_PATTERNS

    # 月別能力を正規化
    def normalize_caps(caps):
        normalized = {}
        for line in DISC_LINES:
            cap = caps.get(line, DEFAULT_CAPACITIES.get(line, 0))
            if isinstance(cap, list):
                normalized[line] = (cap + [cap[-1]] * 12)[:12] if cap else [0] * 12
            else:
                normalized[line] = [cap] * 12
        return normalized

    monthly_capacities = normalize_caps(capacities)

    # 各パターンで最適化実行
    pattern_results = {}
    for rate in patterns:
        try:
            result = optimize(specs, demands, capacities, request.time_limit, load_rate_limit=rate)
            pattern_results[rate] = result
        except Exception as e:
            pattern_results[rate] = None

    # === パターン比較サマリー（2次元配列） ===
    summary_array = [["負荷率上限", "ステータス", "目的関数値", "実行時間(秒)", "平均負荷率", "未割当合計"]]
    for rate in patterns:
        result = pattern_results[rate]
        if result is None:
            summary_array.append([f"{int(rate * 100)}%", "ERROR", "", "", "", ""])
            continue
        total_load = sum(sum(loads) for loads in result.line_loads.values())
        total_cap_annual = sum(sum(c) for c in monthly_capacities.values())
        avg_rate_val = total_load / total_cap_annual if total_cap_annual > 0 else 0
        total_unmet = sum(sum(u) for u in result.unmet_demand.values()) if result.unmet_demand else 0
        summary_array.append([
            f"{int(rate * 100)}%",
            result.status,
            result.objective_value,
            round(result.solve_time, 2),
            f"{avg_rate_val:.1%}",
            total_unmet,
        ])

    # === ライン別負荷率比較（2次元配列） ===
    line_comparison_header = ["ライン", "平均能力"]
    for rate in patterns:
        pct = int(rate * 100)
        line_comparison_header.extend([f"平均負荷({pct}%)", f"負荷率({pct}%)"])
    line_comparison_array = [line_comparison_header]

    for line in DISC_LINES:
        line_caps = monthly_capacities.get(line, [0] * 12)
        avg_cap = sum(line_caps) / 12
        row = [line, int(avg_cap)]
        for rate in patterns:
            result = pattern_results[rate]
            if result is None:
                row.extend(["", ""])
                continue
            loads = result.line_loads.get(line, [0] * 12)
            avg_load = sum(loads) / 12
            load_rate_val = avg_load / avg_cap if avg_cap > 0 else 0
            row.extend([int(avg_load), f"{load_rate_val:.1%}"])
        line_comparison_array.append(row)

    # === ライン別月別負荷（パターン別、2次元配列） ===
    # 各パターンのライン負荷を個別に返す
    patterns_line_loads = {}
    for rate in patterns:
        pct = int(rate * 100)
        result = pattern_results[rate]
        if result is None:
            patterns_line_loads[f"{pct}pct"] = []
            continue

        line_loads_array = [["ライン"] + MONTHS + ["平均能力", "平均負荷", "負荷率"]]
        for line in DISC_LINES:
            loads = result.line_loads.get(line, [0] * 12)
            line_caps = monthly_capacities.get(line, [0] * 12)
            avg_cap = sum(line_caps) / 12
            avg_load = sum(loads) / 12
            load_rate_val = avg_load / avg_cap if avg_cap > 0 else 0
            line_loads_array.append(
                [line] + loads + [int(avg_cap), int(avg_load), f"{load_rate_val:.1%}"]
            )
        patterns_line_loads[f"{pct}pct"] = line_loads_array

    # === 部品割当（パターン別） ===
    patterns_allocations = {}
    for rate in patterns:
        pct = int(rate * 100)
        result = pattern_results[rate]
        if result is None:
            patterns_allocations[f"{pct}pct"] = []
            continue

        alloc_array = [["部品番号", "割当ライン"] + MONTHS + ["年間計"]]
        for part_num in sorted(result.allocation.keys()):
            for line, monthly in result.allocation[part_num].items():
                if sum(monthly) > 0:
                    alloc_array.append([part_num, line] + monthly + [sum(monthly)])
        patterns_allocations[f"{pct}pct"] = alloc_array

    # === 未割当比較（2次元配列） ===
    unmet_comparison_header = ["部品番号"]
    for rate in patterns:
        unmet_comparison_header.append(f"未割当({int(rate * 100)}%)")
    unmet_comparison_array = [unmet_comparison_header]

    all_unmet_parts = set()
    for result in pattern_results.values():
        if result and result.unmet_demand:
            for part_num, monthly in result.unmet_demand.items():
                if sum(monthly) > 0:
                    all_unmet_parts.add(part_num)

    for part_num in sorted(all_unmet_parts):
        row = [part_num]
        for rate in patterns:
            result = pattern_results[rate]
            if result and result.unmet_demand and part_num in result.unmet_demand:
                row.append(sum(result.unmet_demand[part_num]))
            else:
                row.append(0)
        unmet_comparison_array.append(row)

    # === パターン別未割当詳細 ===
    patterns_unmet = {}
    for rate in patterns:
        pct = int(rate * 100)
        result = pattern_results[rate]
        if result is None:
            patterns_unmet[f"{pct}pct"] = []
            continue

        unmet_array = [["部品番号"] + MONTHS + ["年間計"]]
        if result.unmet_demand:
            for part_num in sorted(result.unmet_demand.keys()):
                monthly_unmet = result.unmet_demand[part_num]
                if sum(monthly_unmet) > 0:
                    unmet_array.append([part_num] + monthly_unmet + [sum(monthly_unmet)])
        patterns_unmet[f"{pct}pct"] = unmet_array

    # 全体サマリー
    total_demand = sum(sum(d.monthly_demand) for d in demands.values())
    first_result = pattern_results.get(patterns[0])

    return {
        "success": any(
            r is not None and r.status in ('OPTIMAL', 'FEASIBLE')
            for r in pattern_results.values()
        ),
        "patterns": [int(r * 100) for r in patterns],
        "parts_count": len(specs),
        "total_demand": total_demand,
        # 比較データ
        "comparison_summary": summary_array,
        "line_comparison": line_comparison_array,
        "unmet_comparison": unmet_comparison_array,
        # パターン別詳細データ
        "patterns_line_loads": patterns_line_loads,
        "patterns_allocations": patterns_allocations,
        "patterns_unmet": patterns_unmet,
        # キャパシティ（共通）
        "capacities": [["ライン"] + MONTHS]
            + [[line] + monthly_capacities[line] for line in DISC_LINES],
    }


# ============================================================
# 勤務体制パターン比較API
# ============================================================

class WorkPatternInput(BaseModel):
    """勤務体制パターン入力"""
    name: str                   # 例: "2直2交替"
    formula: str                # 例: "{月間稼働日数} * 7.5 * 2 - {月除外時間}"
    exclusion_hours: float = 0  # 月除外時間


class CompareByWorkPatternRequest(BaseModel):
    """勤務体制パターン比較リクエスト（2次元配列形式）"""
    parts_data: list[list[Any]]
    jph_data: list[list[Any]] | None = None           # [[ライン, JPH], ...]
    work_patterns: list[WorkPatternInput] | None = None
    monthly_working_days: list[float] | None = None   # 12ヶ月分
    time_limit: int = 60


@app.post("/optimize/simple/compare-patterns")
def run_work_pattern_comparison(request: CompareByWorkPatternRequest):
    """
    勤務体制パターンごとに能力を計算し、最適化を実行して比較結果を返す

    jph_data形式: [["ライン", "JPH"], ["4915", 350], ...]
    work_patterns形式: [{"name": "2直2交替", "formula": "...", "exclusion_hours": 5}, ...]
    monthly_working_days形式: [20, 19, 21, ...]  (12ヶ月分)
    """
    # parts_data パース（既存の _parse_simple_request を再利用）
    simple_req = SimpleOptimizeRequest(
        parts_data=request.parts_data,
        capacities_data=None,
        time_limit=request.time_limit,
    )
    specs, demands, _ = _parse_simple_request(simple_req)

    # JPH パース
    jph: dict[str, float] = dict(DEFAULT_JPH)
    if request.jph_data:
        for row in request.jph_data:
            if len(row) < 2:
                continue
            line = str(row[0]).strip()
            if line not in DISC_LINES or line == 'ライン':
                continue
            try:
                jph[line] = float(row[1]) if row[1] else 0
            except (ValueError, TypeError):
                pass

    # 勤務体制パターン パース
    patterns: list[WorkPattern] = []
    if request.work_patterns:
        for wp in request.work_patterns:
            patterns.append(WorkPattern(
                name=wp.name,
                formula=wp.formula,
                exclusion_hours=wp.exclusion_hours,
            ))
    else:
        patterns = list(DEFAULT_WORK_PATTERNS)

    # 月間稼働日数
    working_days = request.monthly_working_days or list(DEFAULT_MONTHLY_WORKING_DAYS)
    if len(working_days) < 12:
        working_days.extend([20] * (12 - len(working_days)))

    # パターン別能力計算
    pattern_capacities = calculate_monthly_capacities(jph, patterns, working_days)

    # 各パターンで最適化実行
    pattern_results = {}
    for pattern_name, capacities in pattern_capacities.items():
        try:
            result = optimize(specs, demands, capacities, request.time_limit)
            pattern_results[pattern_name] = result
        except Exception:
            pattern_results[pattern_name] = None

    pattern_names = [p.name for p in patterns]

    # === パターン比較サマリー ===
    summary_array = [["勤務体制", "ステータス", "目的関数値", "実行時間(秒)", "平均負荷率", "未割当合計"]]
    for name in pattern_names:
        result = pattern_results.get(name)
        capacities = pattern_capacities[name]
        if result is None:
            summary_array.append([name, "ERROR", "", "", "", ""])
            continue
        total_load = sum(sum(loads) for loads in result.line_loads.values())
        total_cap = sum(sum(capacities.get(line, [0] * 12)) for line in DISC_LINES)
        avg_rate_val = total_load / total_cap if total_cap > 0 else 0
        total_unmet = sum(sum(u) for u in result.unmet_demand.values()) if result.unmet_demand else 0
        summary_array.append([
            name,
            result.status,
            result.objective_value,
            round(result.solve_time, 2),
            f"{avg_rate_val:.1%}",
            total_unmet,
        ])

    # === ライン別負荷率比較 ===
    line_comparison_header = ["ライン", "JPH"]
    for name in pattern_names:
        line_comparison_header.extend([f"平均能力({name})", f"平均負荷({name})", f"負荷率({name})"])
    line_comparison_array = [line_comparison_header]

    for line in DISC_LINES:
        row = [line, jph.get(line, 0)]
        for name in pattern_names:
            capacities = pattern_capacities[name]
            result = pattern_results.get(name)
            line_caps = capacities.get(line, [0] * 12)
            avg_cap = sum(line_caps) / 12
            if result is None:
                row.extend(["", "", ""])
                continue
            loads = result.line_loads.get(line, [0] * 12)
            avg_load = sum(loads) / 12
            load_rate_val = avg_load / avg_cap if avg_cap > 0 else 0
            row.extend([int(avg_cap), int(avg_load), f"{load_rate_val:.1%}"])
        line_comparison_array.append(row)

    # === パターン別ライン月別負荷 ===
    patterns_line_loads = {}
    for name in pattern_names:
        capacities = pattern_capacities[name]
        result = pattern_results.get(name)
        if result is None:
            patterns_line_loads[name] = []
            continue

        line_loads_array = [["ライン"] + MONTHS + ["平均能力", "平均負荷", "負荷率"]]
        for line in DISC_LINES:
            loads = result.line_loads.get(line, [0] * 12)
            line_caps = capacities.get(line, [0] * 12)
            avg_cap = sum(line_caps) / 12
            avg_load = sum(loads) / 12
            load_rate_val = avg_load / avg_cap if avg_cap > 0 else 0
            line_loads_array.append(
                [line] + loads + [int(avg_cap), int(avg_load), f"{load_rate_val:.1%}"]
            )
        patterns_line_loads[name] = line_loads_array

    # === パターン別部品割当 ===
    patterns_allocations = {}
    for name in pattern_names:
        result = pattern_results.get(name)
        if result is None:
            patterns_allocations[name] = []
            continue
        alloc_array = [["部品番号", "割当ライン"] + MONTHS + ["年間計"]]
        for part_num in sorted(result.allocation.keys()):
            for line, monthly in result.allocation[part_num].items():
                if sum(monthly) > 0:
                    alloc_array.append([part_num, line] + monthly + [sum(monthly)])
        patterns_allocations[name] = alloc_array

    # === パターン別未割当 ===
    patterns_unmet = {}
    for name in pattern_names:
        result = pattern_results.get(name)
        if result is None:
            patterns_unmet[name] = []
            continue
        unmet_array = [["部品番号"] + MONTHS + ["年間計"]]
        if result.unmet_demand:
            for part_num in sorted(result.unmet_demand.keys()):
                monthly_unmet = result.unmet_demand[part_num]
                if sum(monthly_unmet) > 0:
                    unmet_array.append([part_num] + monthly_unmet + [sum(monthly_unmet)])
        patterns_unmet[name] = unmet_array

    # === 未割当比較 ===
    unmet_comparison_header = ["部品番号"]
    for name in pattern_names:
        unmet_comparison_header.append(f"未割当({name})")
    unmet_comparison_array = [unmet_comparison_header]

    all_unmet_parts = set()
    for result in pattern_results.values():
        if result and result.unmet_demand:
            for part_num, monthly in result.unmet_demand.items():
                if sum(monthly) > 0:
                    all_unmet_parts.add(part_num)

    for part_num in sorted(all_unmet_parts):
        row = [part_num]
        for name in pattern_names:
            result = pattern_results.get(name)
            if result and result.unmet_demand and part_num in result.unmet_demand:
                row.append(sum(result.unmet_demand[part_num]))
            else:
                row.append(0)
        unmet_comparison_array.append(row)

    # === パターン別キャパシティ ===
    patterns_capacities_output = {}
    for name in pattern_names:
        capacities = pattern_capacities[name]
        cap_array = [["ライン"] + MONTHS]
        for line in DISC_LINES:
            cap_array.append([line] + capacities.get(line, [0] * 12))
        patterns_capacities_output[name] = cap_array

    total_demand = sum(sum(d.monthly_demand) for d in demands.values())

    return {
        "success": any(
            r is not None and r.status in ('OPTIMAL', 'FEASIBLE')
            for r in pattern_results.values()
        ),
        "pattern_names": pattern_names,
        "parts_count": len(specs),
        "total_demand": total_demand,
        # 比較データ
        "comparison_summary": summary_array,
        "line_comparison": line_comparison_array,
        "unmet_comparison": unmet_comparison_array,
        # パターン別詳細データ
        "patterns_line_loads": patterns_line_loads,
        "patterns_allocations": patterns_allocations,
        "patterns_unmet": patterns_unmet,
        # パターン別キャパシティ
        "patterns_capacities": patterns_capacities_output,
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
