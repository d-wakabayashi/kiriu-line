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

from config import DISC_LINES, DEFAULT_CAPACITIES, MONTHS
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
        result = optimize(specs, demands, capacities, request.time_limit)
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
    specs = {}
    demands = {}

    # ヘッダー行をスキップして処理
    for row in request.parts_data:
        if len(row) < 16:  # 部品番号 + ライン3つ + 12ヶ月
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

        specs[part_num] = PartSpec(
            part_number=part_num,
            part_name='',
            main_line=main_line,
            sub1_line=sub1_line if sub1_line in DISC_LINES else None,
            sub2_line=sub2_line if sub2_line in DISC_LINES else None,
        )

        demands[part_num] = PartDemand(
            part_number=part_num,
            part_name='',
            monthly_demand=monthly,
        )

    if not specs:
        raise HTTPException(status_code=400, detail="有効な部品データがありません")

    # 能力設定（月別対応）
    capacities = {}
    if request.capacities_data:
        for row in request.capacities_data:
            if len(row) < 2:
                continue
            line = str(row[0]).strip()
            if line not in DISC_LINES or line == 'ライン':
                continue

            if len(row) >= 13:
                # 月別能力形式: [ライン, 4月, 5月, ..., 3月]
                monthly_caps = []
                for i in range(1, 13):
                    try:
                        val = int(float(row[i])) if row[i] else 0
                    except (ValueError, TypeError):
                        val = DEFAULT_CAPACITIES.get(line, 50000)
                    monthly_caps.append(max(0, val))
                capacities[line] = monthly_caps
            else:
                # 固定能力形式: [ライン, 能力]
                try:
                    cap = int(float(row[1]))
                    capacities[line] = cap
                except (ValueError, TypeError):
                    pass

    # 不足しているラインのデフォルト能力を追加
    for line in DISC_LINES:
        if line not in capacities:
            capacities[line] = DEFAULT_CAPACITIES.get(line, 50000)

    # 最適化実行
    result = optimize(specs, demands, capacities, request.time_limit)

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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8080)
