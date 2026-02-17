"""
KIRIU ライン負荷最適化システム - 最適化モデル
"""

from dataclasses import dataclass

from ortools.sat.python import cp_model

from config import (
    DISC_LINES,
    DEFAULT_CAPACITIES,
    WEIGHT_SUB_USE,
    WEIGHT_SUB_QTY,
    DEFAULT_TIME_LIMIT_SECONDS,
    MONTHS,
)
from data_loader import PartSpec, PartDemand


@dataclass
class OptimizationResult:
    """最適化結果"""
    status: str                           # ソルバーステータス
    objective_value: float | None         # 目的関数値
    allocation: dict                      # 部品別ライン別月別割当: {part: {line: [月別数量]}}
    line_loads: dict                      # ライン別月別負荷: {line: [月別負荷]}
    overflow: dict                        # ライン別月別オーバーフロー: {line: [月別超過量]} ※互換性のため残す（常に0）
    sub_line_usage: dict                  # 部品別サブライン使用状況: {part: [月別使用ライン]}
    solve_time: float                     # ソルバー実行時間（秒）
    unmet_demand: dict | None = None      # 部品別月別未割当: {part: [月別未割当量]}


class LineOptimizer:
    """ライン負荷最適化モデル"""

    def __init__(
        self,
        specs: dict[str, PartSpec],
        demands: dict[str, PartDemand],
        capacities: dict[str, int | list[int]] | None = None,
        time_limit: int = DEFAULT_TIME_LIMIT_SECONDS,
        load_rate_limit: float = 1.0,
    ):
        self.specs = specs
        self.demands = demands
        self.time_limit = time_limit
        self.load_rate_limit = load_rate_limit

        # 月別能力に正規化: {ライン: [12ヶ月分の能力]}
        self.capacities = self._normalize_capacities(capacities or DEFAULT_CAPACITIES.copy())

        self.model = cp_model.CpModel()
        self.solver = cp_model.CpSolver()

        # 決定変数格納用
        self.x = {}            # x[part, line, month] = 生産数量
        self.use_sub = {}      # use_sub[part, month] = サブライン使用フラグ
        # 注: オーバーフローは許容しない（ハード制約）

    def _normalize_capacities(self, capacities: dict[str, int | list[int]]) -> dict[str, list[int]]:
        """能力を月別形式に正規化"""
        result = {}
        for line in DISC_LINES:
            cap = capacities.get(line, DEFAULT_CAPACITIES.get(line, 0))
            if isinstance(cap, list):
                if len(cap) == 12:
                    result[line] = cap
                else:
                    # 12ヶ月に満たない場合は最後の値で埋める
                    result[line] = (cap + [cap[-1]] * 12)[:12]
            else:
                result[line] = [cap] * 12
        return result

    def get_capacity(self, line: str, month: int) -> int:
        """指定ライン・月の能力を取得"""
        return self.capacities.get(line, [0] * 12)[month]

    def _get_eligible_lines(self, spec: PartSpec) -> list[str]:
        """部品の割当可能ラインリストを取得"""
        lines = []
        if spec.main_line:
            lines.append(spec.main_line)
        if spec.sub1_line and spec.sub1_line not in lines:
            lines.append(spec.sub1_line)
        if spec.sub2_line and spec.sub2_line not in lines:
            lines.append(spec.sub2_line)
        return lines

    def build_model(self) -> None:
        """最適化モデルを構築"""
        print("最適化モデル構築中...")

        # 大きな数（Big-M）- 月別能力の最大値を使用
        max_cap = max(max(caps) for caps in self.capacities.values())
        M = max_cap * 10

        # ===== 決定変数 =====
        print("  変数作成中...")

        # 部品別ライン別月別生産数量
        for part_num, demand in self.demands.items():
            spec = self.specs.get(part_num)
            if spec is None:
                continue

            eligible_lines = self._get_eligible_lines(spec)
            if not eligible_lines:
                print(f"  警告: {part_num} に割当可能ラインがありません")
                continue

            max_demand = max(demand.monthly_demand) if demand.monthly_demand else 0

            for line in eligible_lines:
                for month in range(12):
                    self.x[part_num, line, month] = self.model.NewIntVar(
                        0, max_demand, f'x_{part_num}_{line}_{month}'
                    )

            # サブライン使用フラグ（メインラインがある場合のみ）
            if spec.main_line and len(eligible_lines) > 1:
                for month in range(12):
                    self.use_sub[part_num, month] = self.model.NewBoolVar(
                        f'use_sub_{part_num}_{month}'
                    )

        # 注: オーバーフローは許容しない（ハード制約）のでオーバーフロー変数は作成しない

        # ===== 制約条件 =====
        print("  制約条件追加中...")

        # 需要未充足変数（ソフト制約用）
        self.unmet_demand = {}

        # 制約1: 需要充足（ソフト制約：未充足を許容）
        for part_num, demand in self.demands.items():
            spec = self.specs.get(part_num)
            if spec is None:
                continue

            eligible_lines = self._get_eligible_lines(spec)
            if not eligible_lines:
                continue

            for month in range(12):
                month_demand = demand.monthly_demand[month]

                # 未充足変数を作成
                self.unmet_demand[part_num, month] = self.model.NewIntVar(
                    0, month_demand, f'unmet_{part_num}_{month}'
                )

                # 該当月の生産量合計 + 未充足 == 需要
                prod_vars = [
                    self.x[part_num, line, month]
                    for line in eligible_lines
                    if (part_num, line, month) in self.x
                ]
                if prod_vars:
                    self.model.Add(
                        sum(prod_vars) + self.unmet_demand[part_num, month] == month_demand
                    )

        # 制約2: ライン能力制約（ハード制約：負荷率上限、オーバーフロー禁止）
        for line in DISC_LINES:
            for month in range(12):
                # 月別の能力を取得し、負荷率上限を適用
                capacity = int(self.get_capacity(line, month) * self.load_rate_limit)

                # このラインの月間総生産量
                line_prod = []
                for part_num in self.demands:
                    if (part_num, line, month) in self.x:
                        line_prod.append(self.x[part_num, line, month])

                if line_prod:
                    # 生産量 <= 能力×負荷率上限（ハード制約：オーバーフロー禁止）
                    self.model.Add(sum(line_prod) <= capacity)

        # 制約3: メインライン優先（サブラインはメインが能力超過の場合のみ）
        for part_num, demand in self.demands.items():
            spec = self.specs.get(part_num)
            if spec is None or spec.main_line is None:
                continue

            eligible_lines = self._get_eligible_lines(spec)
            sub_lines = [l for l in eligible_lines if l != spec.main_line]

            if not sub_lines:
                continue

            for month in range(12):
                # サブライン生産量の合計
                sub_prod = []
                for sub_line in sub_lines:
                    if (part_num, sub_line, month) in self.x:
                        sub_prod.append(self.x[part_num, sub_line, month])

                if not sub_prod:
                    continue

                # サブライン使用フラグとの連動
                if (part_num, month) in self.use_sub:
                    # サブライン生産 > 0 → use_sub = 1
                    # use_sub = 0 → サブライン生産 = 0
                    for sub_line in sub_lines:
                        if (part_num, sub_line, month) in self.x:
                            self.model.Add(
                                self.x[part_num, sub_line, month] <= M * self.use_sub[part_num, month]
                            )

        # ===== 目的関数 =====
        print("  目的関数設定中...")

        objective_terms = []

        # 最優先: 需要未充足最小化（能力制約により生産できなかった分）
        WEIGHT_UNMET = 100000
        for (part_num, month), var in self.unmet_demand.items():
            objective_terms.append(WEIGHT_UNMET * var)

        # 注: オーバーフローは禁止のため、目的関数に含めない

        # 第1優先: サブライン使用回数最小化
        for (part_num, month), var in self.use_sub.items():
            objective_terms.append(WEIGHT_SUB_USE * var)

        # 第2優先: サブラインへの生産量最小化
        for part_num in self.demands:
            spec = self.specs.get(part_num)
            if spec is None or spec.main_line is None:
                continue

            sub_lines = [spec.sub1_line, spec.sub2_line]
            for sub_line in sub_lines:
                if sub_line:
                    for month in range(12):
                        if (part_num, sub_line, month) in self.x:
                            objective_terms.append(
                                WEIGHT_SUB_QTY * self.x[part_num, sub_line, month]
                            )

        self.model.Minimize(sum(objective_terms))

        print(f"  モデル構築完了: 変数数={len(self.x) + len(self.use_sub)}")

    def solve(self) -> OptimizationResult:
        """最適化を実行"""
        print(f"\n最適化実行中（制限時間: {self.time_limit}秒）...")

        self.solver.parameters.max_time_in_seconds = self.time_limit
        self.solver.parameters.num_search_workers = 8  # 並列探索

        status = self.solver.Solve(self.model)
        solve_time = self.solver.WallTime()

        status_names = {
            cp_model.OPTIMAL: 'OPTIMAL',
            cp_model.FEASIBLE: 'FEASIBLE',
            cp_model.INFEASIBLE: 'INFEASIBLE',
            cp_model.MODEL_INVALID: 'MODEL_INVALID',
            cp_model.UNKNOWN: 'UNKNOWN',
        }
        status_str = status_names.get(status, f'UNKNOWN({status})')
        print(f"  ステータス: {status_str}")
        print(f"  実行時間: {solve_time:.2f}秒")

        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return OptimizationResult(
                status=status_str,
                objective_value=None,
                allocation={},
                line_loads={},
                overflow={},
                sub_line_usage={},
                solve_time=solve_time,
            )

        obj_value = self.solver.ObjectiveValue()
        print(f"  目的関数値: {obj_value:,.0f}")

        # 需要未充足を確認
        total_unmet = 0
        unmet_parts = []
        for (part_num, month), var in self.unmet_demand.items():
            unmet_val = self.solver.Value(var)
            if unmet_val > 0:
                total_unmet += unmet_val
                unmet_parts.append((part_num, month, unmet_val))

        if total_unmet > 0:
            print(f"  警告: 需要未充足あり - 合計 {total_unmet:,}")
            for part, month, qty in unmet_parts[:5]:
                print(f"    {part} / {month+1}月: {qty:,}")

        # 結果を抽出
        allocation = {}
        for part_num in self.demands:
            spec = self.specs.get(part_num)
            if spec is None:
                continue

            eligible_lines = self._get_eligible_lines(spec)
            allocation[part_num] = {}

            for line in eligible_lines:
                monthly = []
                for month in range(12):
                    if (part_num, line, month) in self.x:
                        val = self.solver.Value(self.x[part_num, line, month])
                    else:
                        val = 0
                    monthly.append(val)

                if sum(monthly) > 0:
                    allocation[part_num][line] = monthly

        # ライン別負荷を集計
        line_loads = {line: [0] * 12 for line in DISC_LINES}
        for part_num, lines in allocation.items():
            for line, monthly in lines.items():
                for month, qty in enumerate(monthly):
                    line_loads[line][month] += qty

        # オーバーフロー量（ハード制約により常に0）
        overflow = {line: [0] * 12 for line in DISC_LINES}

        # 未割当（需要未充足）を部品別月別で抽出
        unmet_by_part = {}
        for part_num in self.demands:
            monthly_unmet = []
            for month in range(12):
                if (part_num, month) in self.unmet_demand:
                    unmet_val = self.solver.Value(self.unmet_demand[part_num, month])
                else:
                    unmet_val = 0
                monthly_unmet.append(unmet_val)
            if sum(monthly_unmet) > 0:
                unmet_by_part[part_num] = monthly_unmet

        # サブライン使用状況
        sub_usage = {}
        for part_num in self.demands:
            spec = self.specs.get(part_num)
            if spec is None:
                continue

            sub_usage[part_num] = []
            for month in range(12):
                used_lines = []
                for line, monthly in allocation.get(part_num, {}).items():
                    if monthly[month] > 0:
                        if line == spec.main_line:
                            used_lines.insert(0, line)
                        else:
                            used_lines.append(line)
                sub_usage[part_num].append(used_lines)

        return OptimizationResult(
            status=status_str,
            objective_value=obj_value,
            allocation=allocation,
            line_loads=line_loads,
            overflow=overflow,
            sub_line_usage=sub_usage,
            solve_time=solve_time,
            unmet_demand=unmet_by_part,
        )


def optimize(
    specs: dict[str, PartSpec],
    demands: dict[str, PartDemand],
    capacities: dict[str, int | list[int]] | None = None,
    time_limit: int = DEFAULT_TIME_LIMIT_SECONDS,
    load_rate_limit: float = 1.0,
) -> OptimizationResult:
    """
    最適化を実行するヘルパー関数

    Args:
        specs: 部品仕様辞書
        demands: 部品需要辞書
        capacities: ライン能力辞書（オプション）
        time_limit: ソルバー制限時間（秒）
        load_rate_limit: 負荷率上限（0.0〜1.0、デフォルト1.0=100%）

    Returns:
        OptimizationResult
    """
    optimizer = LineOptimizer(specs, demands, capacities, time_limit, load_rate_limit)
    optimizer.build_model()
    return optimizer.solve()


if __name__ == '__main__':
    # テスト実行
    from data_loader import load_all_data

    specs, demands = load_all_data()
    result = optimize(specs, demands, time_limit=60)

    print("\n=== 最適化結果サマリー ===")
    print(f"ステータス: {result.status}")

    if result.allocation:
        print("\nライン別月間負荷:")
        for line in DISC_LINES:
            loads = result.line_loads.get(line, [0] * 12)
            cap = DEFAULT_CAPACITIES.get(line, 0)
            avg_load = sum(loads) / 12
            avg_rate = avg_load / cap * 100 if cap > 0 else 0
            print(f"  {line}: 平均 {avg_load:,.0f} ({avg_rate:.1f}%)")

        # オーバーフロー確認
        total_overflow = sum(sum(v) for v in result.overflow.values())
        if total_overflow > 0:
            print(f"\n総オーバーフロー: {total_overflow:,}")
            for line in DISC_LINES:
                line_of = result.overflow.get(line, [0] * 12)
                if sum(line_of) > 0:
                    print(f"  {line}: {line_of}")
