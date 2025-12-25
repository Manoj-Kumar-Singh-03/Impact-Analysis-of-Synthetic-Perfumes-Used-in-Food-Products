import pandas as pd
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import seaborn as sns

# ───────────────── DATA  +  CLEANING ─────────────────
df = pd.read_excel("perfumedata.xlsx")
df.columns = df.columns.str.strip()
df["Causes/Effects"] = (
    df["Causes/Effects"].str.lower().replace(
        {
            "irritation effects":  "Irritation Effects",
            "specialized effects": "Specialized Effects",
            "specified effects":   "Specialized Effects",
            "systemic toxicity":   "Systemic Toxicity",
            "safety approvals":    "Safety Approvals",
            "allergic reaction":   "Allergic Reaction",
        }
    )
)

# ───────────────── UNIVERSAL PLOTTERS ────────────────
def heatmap(table, title, cmap="YlGnBu"):
    fig, ax = plt.subplots(figsize=(12, 7))
    sns.heatmap(table, annot=True, fmt=".0f", cmap=cmap, ax=ax)
    ax.set_title(title); ax.set_xlabel(""); ax.set_ylabel("")
    fig.tight_layout(); return fig

def correlation(table, title):
    fig, ax = plt.subplots(figsize=(8, 6))
    sns.heatmap(table.corr(), annot=True, cmap="coolwarm", ax=ax)
    ax.set_title(title); fig.tight_layout(); return fig

def pie(series, title):
    fig, ax = plt.subplots(figsize=(6, 6))
    series.value_counts().plot(
        kind="pie", autopct="%1.1f%%", startangle=90,
        ax=ax, colors=plt.cm.Paired.colors
    )
    ax.set_title(title); ax.set_ylabel("")
    fig.tight_layout(); return fig

def bar(series, title, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(8, 4))
    series.value_counts().plot(kind="bar", ax=ax)
    ax.set_title(title); ax.set_xlabel(xlabel); ax.set_ylabel(ylabel)
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout(); return fig

def scatter(x, y, title, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.scatter(x, y); ax.set_title(title)
    ax.set_xlabel(xlabel); ax.set_ylabel(ylabel)
    fig.tight_layout(); return fig

def box(series, title, ylabel):
    fig, ax = plt.subplots(figsize=(6, 4))
    sns.boxplot(y=series, ax=ax)
    ax.set_title(title); ax.set_ylabel(ylabel)
    fig.tight_layout(); return fig

# ─────────────── QUESTION  →  CHART BUILDERS ─────────
def build_distribution():
    s = df["Causes/Effects"]
    return {
        "bar": bar(s, "Distribution of Effects", "Effect", "Count"),
        "pie": pie(s, "Effects – share"),
        "box": box(s.value_counts(), "Spread of Effect Counts", "Count"),
    }

def build_flavour():
    df_flav = df.copy()
    if "Chemical Formula" in df_flav.columns:
        df_flav["Flavour of Aroma"] = df_flav.apply(
            lambda x: f"{x['Flavour of Aroma']} ({x['Chemical Formula']})", axis=1)
    pivot = df_flav.groupby(["Flavour of Aroma", "Causes/Effects"]).size().unstack(fill_value=0)
    return {
        "heat": heatmap(pivot, "Effects by Flavour"),
        "corr": correlation(pivot, "Correlation between Effects (Flavour table)"),
    }

def build_food():
    pivot = df.groupby(["Food", "Causes/Effects"]).size().unstack(fill_value=0)
    return {
        "heat": heatmap(pivot, "Effects by Food", cmap="YlOrRd"),
        "corr": correlation(pivot, "Correlation between Effects (Food table)"),
    }

def build_slice(effect):
    sub = df[df["Causes/Effects"] == effect].copy()
    if "Chemical Formula" in sub.columns:
        sub["Flavour of Aroma"] = sub.apply(
            lambda x: f"{x['Flavour of Aroma']} ({x['Chemical Formula']})", axis=1)
    return {
        "pie": pie(sub["Flavour of Aroma"], f"Flavours causing {effect}"),
        "bar": bar(sub["Food"], f"Food count – {effect}", "Food", "Count"),
        "scatter": scatter(range(len(sub)), sub.index,
                           f"Index scatter – {effect}", "Index", "Row #"),
    }

question_map = {
    "Distribution of Effects": build_distribution,
    "Effects by Flavour of Aroma": build_flavour,
    "Effects by Food Type": build_food,
}
for eff in df["Causes/Effects"].unique():
    question_map[f"{eff} (specific)"] = lambda eff=eff: build_slice(eff)

# ────────────────── TKINTER GUI ──────────────────────
root = tk.Tk(); root.title("Interactive Perfume Analytics")
root.geometry("1150x780")

# Dropdown
ttk.Label(root, text="Select Question:", font=("Segoe UI", 11))\
    .pack(anchor="w", pady=(8, 0), padx=8)
question_cb = ttk.Combobox(root, values=list(question_map), state="readonly", width=45)
question_cb.pack(anchor="w", padx=8); question_cb.current(0)

# Frame that will hold (and refresh) check‑buttons
style_frame = ttk.Frame(root); style_frame.pack(anchor="w", padx=8, pady=6)
style_vars = {}     # key → BooleanVar  (rebuilt each refresh)

def rebuild_checkbuttons(event=None):
    # clear existing widgets
    for child in style_frame.winfo_children():
        child.destroy()
    style_vars.clear()

    # build new ones for this question
    available = question_map[question_cb.get()]().keys()
    label_map = {
        "heat": "Heat‑map", "corr": "Correlation", "pie": "Pie chart",
        "bar": "Bar chart", "scatter": "Scatter plot", "box": "Box plot"
    }
    for key in available:
        var = tk.BooleanVar(value=False)
        cb  = ttk.Checkbutton(style_frame, text=label_map.get(key, key), variable=var)
        cb.pack(side=tk.LEFT, padx=4)
        style_vars[key] = var

question_cb.bind("<<ComboboxSelected>>", rebuild_checkbuttons)
rebuild_checkbuttons()   # initial buttons

# Draw button
draw_btn = ttk.Button(root, text="Draw"); draw_btn.pack(anchor="w", padx=8, pady=(0, 8))

# Canvas display
display = tk.Frame(root); display.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
canvases = []
def clear_canvases():
    for c in canvases: c.get_tk_widget().destroy()
    canvases.clear()

def draw():
    clear_canvases()
    figs = question_map[question_cb.get()]()
    for key, fig in figs.items():
        if style_vars[key].get():
            cv = FigureCanvasTkAgg(fig, master=display)
            cv.draw(); cv.get_tk_widget().pack(fill=tk.BOTH, expand=True, pady=4)
            canvases.append(cv)

draw_btn.configure(command=draw)

root.mainloop()
