import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import numpy as np


# --- Step 1: Find latest Excel file ---
logs_folder = Path("Logs")
excel_files = list(logs_folder.glob("*.xlsx"))
if not excel_files:
    raise FileNotFoundError("No Excel files found in Logs folder.")

latest_file = max(excel_files, key=lambda f: f.stat().st_mtime)
print(f"Using latest file: {latest_file.name}")

# --- Step 2: Load data ---
df = pd.read_excel(latest_file, header=1)

# --- Step 3: Split data into runs using Data type 99 as delimiter ---
run_indices = df.index[df['Data type'] == 99].tolist()
runs = []
start_idx = 0
for idx in run_indices:
    run = df.iloc[start_idx:idx]  # exclude the 99 row
    runs.append(run)
    start_idx = idx + 1  # next run starts after 99

# --- Step 4: Calculate statistics ---
run_stats = []
max_volume = df[df['Data type'] == 2]['Data'].iloc[0]

for i, run in enumerate(runs, start=1):
    times = run[run['Data type'] == 6]['Data'].values
    steps = run[run['Data type'] == 4]['Data'].values
    avg_time = np.mean(times) if len(times) > 0 else np.nan
    total_iterations = len(steps)
    convergence_speed = np.mean(-np.diff(steps)) if len(steps) > 1 else 0
    run_stats.append({
        "Run": i,
        "Avg Time": avg_time,
        "Iterations": total_iterations,
        "Convergence Speed": convergence_speed
    })

stats_df = pd.DataFrame(run_stats)

# --- Step 5: Plot Steps Remaining vs Iteration ---
plt.figure(figsize=(12,6))
colors = plt.cm.viridis(np.linspace(0, 1, len(runs)))
for run, color in zip(runs, colors):
    steps_remaining = run[run['Data type'] == 3]['Data'].values
    iterations = np.arange(1, len(steps_remaining)+1)
    plt.plot(iterations, steps_remaining, marker='o', alpha=0.7, color=color, linewidth=1.5)

plt.title(f"Steps Remaining vs Iteration\n(Max Volume = {max_volume})", fontsize=14)
plt.xlabel("Iteration", fontsize=12)
plt.ylabel("Steps Remaining", fontsize=12)
plt.grid(True, linestyle='--', alpha=0.5)
plt.tight_layout()
plt.show()

# --- Step 6: Show statistics in a separate, pretty table window ---
fig, ax = plt.subplots(figsize=(8, len(stats_df)*0.6 + 1))
ax.axis('off')

# Add table
table = ax.table(
    cellText=stats_df.round(2).values.tolist(),
    colLabels=stats_df.columns,
    cellLoc='center',
    loc='center'
)

# Improve aesthetics
table.auto_set_font_size(False)
table.set_fontsize(12)
table.auto_set_column_width(col=list(range(len(stats_df.columns))))

# Header formatting
for (i, j), cell in table.get_celld().items():
    if i == 0:  # header
        cell.set_fontsize(14)
        cell.set_text_props(weight='bold', color='white')
        cell.set_facecolor('#4CAF50')  # green header
    else:
        # Alternate row colors
        if i % 2 == 0:
            cell.set_facecolor('#f1f1f1')
        else:
            cell.set_facecolor('white')

ax.set_title("Run Statistics", fontsize=16, weight='bold')
plt.tight_layout()
plt.show()
