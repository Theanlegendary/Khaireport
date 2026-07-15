import matplotlib.pyplot as plt
import numpy as np

# Set style
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

categories = ['Pickup', 'Delivery', 'Pending']
morning = [49, 149, 748]
afternoon = [108, 232, 684]

# Calculate changes
diff = [a - m for m, a in zip(morning, afternoon)]
pct = [d / m * 100 for m, d in zip(morning, diff)]

x = np.arange(len(categories))
width = 0.35

# 1. Bar Chart (Volume Comparison)
rects1 = ax1.bar(x - width/2, morning, width, label='Morning (9 AM)', color='#C00000') # Red for urgent
rects2 = ax1.bar(x + width/2, afternoon, width, label='Afternoon (2 PM)', color='#1F4E78')

ax1.set_ylabel('Number of Urgent (Overdue) Orders', fontsize=12, fontweight='bold', color='#2F5597')
ax1.set_title('Urgent Orders Volume Comparison (9 AM vs 2 PM)', fontsize=14, fontweight='bold', pad=15)
ax1.set_xticks(x)
ax1.set_xticklabels(categories, fontsize=11, fontweight='bold')
ax1.legend(frameon=True, facecolor='white', edgecolor='none')

# Add values on top of bars
def autolabel(rects, ax):
    for rect in rects:
        height = rect.get_height()
        ax.annotate(f'{height}',
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=10, fontweight='bold')

autolabel(rects1, ax1)
autolabel(rects2, ax1)

# 2. Bar Chart (Percentage Growth)
colors_pct = ['#70AD47' if p > 0 else '#C00000' for p in pct]
rects3 = ax2.bar(categories, pct, width=0.5, color=colors_pct, edgecolor='none')

ax2.set_ylabel('Percentage Change (%)', fontsize=12, fontweight='bold', color='#2F5597')
ax2.set_title('Percentage Change (%) in Urgent Orders', fontsize=14, fontweight='bold', pad=15)

# Dynamic y limit
y_min = min(pct) * 1.2 if min(pct) < 0 else 0
y_max = max(pct) * 1.2 if max(pct) > 0 else 10
ax2.set_ylim(y_min, y_max)

for rect, p, d in zip(rects3, pct, diff):
    height = rect.get_height()
    va_dir = 'bottom' if p >= 0 else 'top'
    xy_offset = (0, 3) if p >= 0 else (0, -15)
    sign = "+" if p >= 0 else ""
    ax2.annotate(f'{sign}{p:.1f}%\n({d:+} orders)',
                xy=(rect.get_x() + rect.get_width() / 2, height if p >= 0 else height),
                xytext=xy_offset,
                textcoords="offset points",
                ha='center', va=va_dir, fontsize=10, fontweight='bold')

plt.tight_layout()

# Save chart to artifact directory
chart_path = r"c:\Users\DELL\.gemini\antigravity\brain\b8105454-e4a4-4a86-b723-6ef7dcbedc27\report_urgent_chart.png"
plt.savefig(chart_path, dpi=300)
print(f"Urgent chart saved to: {chart_path}")
