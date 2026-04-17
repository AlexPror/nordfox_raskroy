"""Интерфейс tkinter (если PySide6 не установлен)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from nordfox_raskroy.excel_io import parse_specification
from nordfox_raskroy.optimizer import (
    format_cut_angles,
    optimize_cutting,
    sort_cuts_for_display,
    spec_rows_to_demands,
    summarize,
)
from nordfox_raskroy import __version__
from nordfox_raskroy.materials_library import row_mass_kg_display
from nordfox_raskroy.profile_codes import PROFILE_DIGIT_TO_NAME, filter_spec_by_profiles, profile_label_for_code


def run_tk_app() -> None:
    root = tk.Tk()
    root.title(f"NordFox — раскрой v{__version__} (tkinter · установите PySide6 для Qt)")
    root.geometry("1180x680")

    path_var = tk.StringVar()
    kerf_var = tk.StringVar(value="0")
    min_scrap_var = tk.StringVar(value="50")

    v6000 = tk.BooleanVar(value=True)
    v7500 = tk.BooleanVar(value=True)
    v12000 = tk.BooleanVar(value=True)

    profile_vars: dict[int, tk.BooleanVar] = {
        i: tk.BooleanVar(value=True) for i in PROFILE_DIGIT_TO_NAME
    }

    top = ttk.Frame(root, padding=8)
    top.pack(fill=tk.X)

    ttk.Label(top, text="Файл спецификации (.xlsx):").grid(row=0, column=0, sticky="w")

    def browse() -> None:
        p = filedialog.askopenfilename(
            title="Спецификация",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if p:
            path_var.set(p)

    ttk.Entry(top, textvariable=path_var, width=72).grid(row=0, column=1, padx=4)
    ttk.Button(top, text="Обзор…", command=browse).grid(row=0, column=2)

    prof_fr = ttk.LabelFrame(
        root,
        text="Профили (СК-/СС-/Р-: 0=Н20 … 3=Н23)",
        padding=8,
    )
    prof_fr.pack(fill=tk.X, padx=8, pady=4)
    pf = ttk.Frame(prof_fr)
    pf.pack(fill=tk.X)
    for i, name in PROFILE_DIGIT_TO_NAME.items():
        ttk.Checkbutton(pf, text=f"{name} ({i})", variable=profile_vars[i]).pack(
            side=tk.LEFT, padx=8
        )

    bar_fr = ttk.LabelFrame(root, text="Длины заготовок (мм)", padding=8)
    bar_fr.pack(fill=tk.X, padx=8, pady=4)
    ttk.Checkbutton(bar_fr, text="6000", variable=v6000).pack(side=tk.LEFT, padx=6)
    ttk.Checkbutton(bar_fr, text="7500", variable=v7500).pack(side=tk.LEFT, padx=6)
    ttk.Checkbutton(bar_fr, text="12000", variable=v12000).pack(side=tk.LEFT, padx=6)

    opt_fr = ttk.Frame(root, padding=8)
    opt_fr.pack(fill=tk.X)
    ttk.Label(opt_fr, text="Пропил (kerf), мм:").pack(side=tk.LEFT)
    ttk.Entry(opt_fr, textvariable=kerf_var, width=6).pack(side=tk.LEFT, padx=6)
    ttk.Label(opt_fr, text="Мин. обрезок, мм:").pack(side=tk.LEFT, padx=(16, 0))
    ttk.Entry(opt_fr, textvariable=min_scrap_var, width=6).pack(side=tk.LEFT, padx=6)

    btn_fr = ttk.Frame(root, padding=8)
    btn_fr.pack(fill=tk.X)

    summary = tk.Text(root, height=6, wrap="word", font=("Segoe UI", 10))
    summary.pack(fill=tk.X, padx=8, pady=4)

    tree_fr = ttk.Frame(root, padding=8)
    tree_fr.pack(fill=tk.BOTH, expand=True)

    cols = (
        "module",
        "profile",
        "series",
        "len",
        "angle",
        "src",
        "stock",
        "rem",
        "mass",
    )
    tree = ttk.Treeview(
        tree_fr,
        columns=cols,
        show="headings",
        height=18,
    )
    headings = {
        "module": "Модуль",
        "profile": "Тип профиля",
        "series": "Серия",
        "len": "Длина",
        "angle": "Угол",
        "src": "Источник",
        "stock": "Заготовка мм",
        "rem": "Остаток мм",
        "mass": "Масса профиля, кг",
    }
    for c in cols:
        tree.heading(c, text=headings[c])
        w = 180 if c == "profile" else 100
        tree.column(c, width=w, anchor="w")
    vsb = ttk.Scrollbar(tree_fr, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def compute() -> None:
        path = path_var.get().strip()
        if not path or not Path(path).is_file():
            messagebox.showerror("Ошибка", "Укажите существующий файл .xlsx")
            return
        bars: list[int] = []
        if v6000.get():
            bars.append(6000)
        if v7500.get():
            bars.append(7500)
        if v12000.get():
            bars.append(12000)
        if not bars:
            messagebox.showerror("Ошибка", "Выберите хотя бы одну длину заготовки")
            return
        allowed = {i for i, v in profile_vars.items() if v.get()}
        if not allowed:
            messagebox.showerror("Ошибка", "Отметьте хотя бы один профиль Н20–Н23")
            return
        try:
            kerf = int(kerf_var.get().strip() or "0")
            min_scrap = int(min_scrap_var.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Ошибка", "Пропил и мин. обрезок должны быть целыми числами")
            return
        try:
            rows_all = parse_specification(path)
            rows, warns = filter_spec_by_profiles(rows_all, allowed)
            if not rows:
                wtxt = "\n".join(warns[:25]) if warns else ""
                messagebox.showwarning(
                    "Нет данных",
                    "После фильтра не осталось строк.\n" + wtxt,
                )
                return
            demands = spec_rows_to_demands(rows)
            result = optimize_cutting(
                demands,
                bar_lengths_mm=bars,
                kerf_mm=kerf,
                min_scrap_mm=min_scrap,
            )
        except Exception as e:  # noqa: BLE001 — прототип
            messagebox.showerror("Расчёт", str(e))
            return

        cuts_sorted = sort_cuts_for_display(result.cuts, by_module=True)
        total_kg = 0.0
        any_mass = False
        for cut in cuts_sorted:
            kg, _ = row_mass_kg_display(
                cut.demand.profile_code, float(cut.demand.length_mm), 1.0
            )
            if kg is not None:
                total_kg += kg
                any_mass = True
        tot_s = (
            f"{total_kg:.3f}".rstrip("0").rstrip(".")
            if any_mass
            else "—"
        )

        summary.delete("1.0", tk.END)
        lines = [summarize(result)]
        lines.append(f"Итого масса профиля, кг: {tot_s if tot_s else '0'}")
        if warns:
            lines.append(f"Исключено при фильтре: {len(warns)}")
            lines.append("\n".join(warns[:30]))
            if len(warns) > 30:
                lines.append(f"… ещё {len(warns) - 30}")
        summary.insert(tk.END, "\n".join(lines))

        if warns:
            messagebox.showinfo(
                "Фильтр",
                f"Исключено строк: {len(warns)}. См. текст сводки выше.",
            )

        for item in tree.get_children():
            tree.delete(item)
        for cut in cuts_sorted:
            d = cut.demand
            _, mtxt = row_mass_kg_display(d.profile_code, float(d.length_mm), 1.0)
            tree.insert(
                "",
                tk.END,
                values=(
                    d.module_name,
                    d.profile_code,
                    profile_label_for_code(d.profile_code),
                    d.length_mm,
                    format_cut_angles(d),
                    "Обрезок" if cut.stock_source == "scrap" else "Новая",
                    cut.stock_length_mm,
                    cut.remainder_mm,
                    mtxt,
                ),
            )

    ttk.Button(btn_fr, text="Рассчитать раскрой", command=compute).pack(side=tk.LEFT)

    default_spec = Path(__file__).resolve().parents[2] / "test" / "spec_10x5_modules.xlsx"
    if default_spec.is_file():
        path_var.set(str(default_spec))

    root.mainloop()


if __name__ == "__main__":
    run_tk_app()
