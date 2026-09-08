# Targetpend Analysis Project

This directory contains the setup for daily data analysis using Jupyter Notebook.

## Directory Structure

- `source/` - Place daily source files here.
- `output/` - The destination directory for generated outputs and reports.
- `targetpend/` - Python virtual environment containing required packages (Jupyter, Pandas, etc.).
- `requirements.txt` - Python package dependencies.
- `.gitignore` - Standard gitignore configurations.

## Quick Start

### 1. Activate the Virtual Environment
Open a terminal in this directory (`F:\targetpend`) and run the activation command for your shell:

- **PowerShell**:
  ```powershell
  .\targetpend\Scripts\Activate.ps1
  ```
- **Command Prompt (CMD)**:
  ```cmd
  .\targetpend\Scripts\activate.bat
  ```

### 2. Start Jupyter Notebook
Once the virtual environment is activated, launch Jupyter:
```bash
jupyter notebook
```
or JupyterLab:
```bash
jupyter lab
```
