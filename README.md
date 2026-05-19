# Fugro JDF Valve Torque Analyzer 📊🔩

A premium, standalone Windows desktop application designed for **Fugro** to automate the batch processing, analysis, and report generation of mechanical valve torque **Job Definition Format (JDF)** log files. 

This program replaces manual spreadsheets and screenshotting routines with an elegant, single-click GUI utility that converts raw data into polished Excel files and vector-perfect PDF engineering reports.

> [!NOTE]
> Engineered and calibrated exclusively for Fugro operations to ensure seamless compliance and precision.


---

## 🌟 Key Features

* **Instant Batch Processing**: Drag-and-drop or batch-load dozens of JDF files simultaneously to analyze them in seconds.
* **Intelligent Torque Calibrator**: Automatic statistical detection of critical mechanical valve torque phases:
  * **Break Out Torque**: Identifies the initial peak torque required to break static friction.
  * **Running Torque**: Computes the steady-state plateau travel torque (calculated via statistical mode/average of the active run).
  * **Make Up Torque**: Captures the final seating torque achieved at the close phase.
* **Full Parameter Control**: Operators can review and manually adjust the calculated parameters in the UI before generation, giving absolute engineering control.
* **Interactive Excel Generation**: Automated writing of clean timestamps and torque columns with embedded formulas (`=MAX(C:C)` and `=MIN(C:C)`) and a native Excel ScatterChart.
* **Vector-Perfect PDF Reports**: Combines Fugro branding, custom metadata, and professional high-resolution inline charts into a highly compressed, A4-ready PDF report (averaging only **~50 KB**!).
* **Zero-Dependency Executable**: The final compiled program (`FugroJDFAnalyzer.exe`) runs on any Windows machine instantly with **no Python installation, libraries, or external assets required**.

---

## 📁 Repository Structure

```
valveTorque/
├── example-assets/             # Original templates, raw JDF logs, and extracted assets
│   ├── VN14090 Fully Close-1.JDF
│   ├── VN14090 Fully Close-1.xlsx
│   ├── VN14090 Fully Close-1.pdf
│   ├── extracted_img_p1_0.png  # Embedded Fugro logo asset
│   └── app_icon.ico            # High-resolution multi-size Windows icon
├── app.py                      # Main PyQt5 Desktop Application (GUI & Workers)
├── jdf_converter_backend.py    # Standalone command-line JDF converter backend
├── plot_jdf_matplotlib.py      # Independent Matplotlib torque-time chart renderer
├── build_exe.bat               # Compiles app.py inside a clean virtual environment
├── jdf_analyzer_implementation_plan.md # Deep-dive technical implementation plan
├── .gitignore                  # Git exclusion rules
└── README.md                   # This professional documentation
```

---

## 🚀 How to Build the Standalone Executable

To compile the application into a lightweight, portable Windows executable (`.exe`) that can run on any machine without Python:

1. Open PowerShell or Command Prompt.
2. Navigate to the repository directory:
   ```bash
   d:
   cd d:\Fugro\valveTorque
   ```
3. Run the automated clean virtual environment build batch file:
   ```bash
   .\build_exe.bat
   ```

*This will automatically create a clean, isolated virtual environment, install only the lightweight target dependencies, and run PyInstaller. The output will be created inside the **`dist`** directory as `dist\FugroJDFAnalyzer.exe` (~35 MB).*

---

## 🖥️ How to Run & Use the Application

### 1. Launch the Application
Simply double-click **`dist\FugroJDFAnalyzer.exe`** (or run `python app.py` from your terminal if developing). 

### 2. Load JDF Log Files
* Click the **`Add JDF Files...`** button.
* Select one or multiple `.JDF` files from your files manager.
* They will instantly load into the left-hand **Loaded JDF Files** queue.

### 3. Review and Customize Torque Parameters
* Click on any file in the queue to load its auto-calculated metrics into the **Selected File Parameters** panel on the right.
* The system automatically populates:
  * **Break Out Torque**
  * **Running Torque**
  * **Make Up Torque**
* If you want to override a metric (e.g. to set a target nominal Make Up torque of `300` Nm instead of the exact logged peak of `313` Nm), simply edit the text box. The edits save automatically.

### 4. Process and Export Reports
* Select your **Output Folder** path.
* Click **`PROCESS ALL JDF FILES`**.
* The application runs the batch conversions in the background (keeping the UI fully fluid and responsive) and updates the progress bar.
* Find your completed professional `.xlsx` sheets and A4-ready `.pdf` reports immediately inside your selected output folder!

---

## 🔧 Technical Details & Calibration

### Torque-Time Scaling Alignment
Excel uses a day fraction scale to plot scatter charts. To guarantee 100% matching visual alignment on the generated PDF charts, our custom plotting engine implements this exact auto-scaling logic:
* **X-Axis Range**: Start: `14:03:50` | End: `14:21:07`
* **X-Axis Intervals**: Step exactly at `02:53` major units (`14:03:50`, `14:06:43`, `14:09:36`, etc.) representing `0.002` day fractions.
* **Y-Axis Range**: Limits locked at `[0, 350]` Nm with gridlines every `50` Nm.

---

## 🛡️ License and Ownership
Developed specifically for **Fugro** operations. All rights reserved.
