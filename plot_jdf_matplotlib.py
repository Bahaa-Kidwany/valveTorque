import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
import numpy as np
from datetime import datetime, time

# Read data from JDF
times = []
torques = []
with open(r"D:\Fugro\valveTorque\example-assets\VN14090 Fully Close-1.JDF", "r") as f:
    for idx, line in enumerate(f):
        parts = line.strip().split("\t")
        if idx >= 144:  # Active data starts at line 145
            t_str = parts[1]
            torque = abs(float(parts[5]))
            
            # Convert timestamp to datetime object for matplotlib
            dt = datetime.strptime(t_str, "%H:%M:%S")
            times.append(dt)
            torques.append(torque)

times = np.array(times)
torques = np.array(torques)

# Create high-resolution figure matching the aspect ratio of the PDF chart
fig, ax = plt.subplots(figsize=(10, 6), dpi=300)

# Set white background
fig.patch.set_facecolor('white')
ax.set_facecolor('white')

# Plot the torque line (blue, matching the spreadsheet chart style)
ax.plot(times, torques, color='#1F77B4', linewidth=1.5, label='Torque')

# Configure axes limits to match Excel auto-scale EXACTLY:
# Min time: 14:03:50
# Max time: 14:21:07
min_time = datetime.strptime("14:03:50", "%H:%M:%S")
max_time = datetime.strptime("14:21:07", "%H:%M:%S")
ax.set_xlim(min_time, max_time)

# Configure Y axis limits and ticks: 0 to 350 with step of 50
ax.set_ylim(0, 350)
ax.set_yticks(range(0, 351, 50))

# Configure X axis format and major ticks: step of 2 minutes 53 seconds
# We can set specific ticks at the matching locations:
ticks = [
    "14:03:50",
    "14:06:43",
    "14:09:36",
    "14:12:29",
    "14:15:22",
    "14:18:14",
    "14:21:07"
]
tick_dts = [datetime.strptime(t, "%H:%M:%S") for t in ticks]
ax.set_xticks(tick_dts)
ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))

# Enable grid lines (light gray, thin)
ax.grid(True, which='both', color='#E0E0E0', linestyle='-', linewidth=0.5)

# Axis labels
ax.set_ylabel('Torque\n- CCW   Nm /  + CW  Nm', fontsize=9, color='#333333')
ax.set_xlabel('Time', fontsize=9, color='#333333')

# Clean chart border
for spine in ax.spines.values():
    spine.set_color('#CCCCCC')
    spine.set_linewidth(0.8)

# Title
ax.set_title('VN14090 Fully Close-1', fontsize=12, fontweight='bold', pad=15)

plt.tight_layout()
plt.savefig(r"D:\Fugro\valveTorque\example-assets\matplotlib_chart_test.png", facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
print("Successfully generated matplotlib_chart_test.png")
