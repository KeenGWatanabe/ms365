# ms365
pull users data from MS Graph
To run your Python code smoothly from the VS Code terminal (especially inside WSL), here’s how to go about it:

---

### 🧭 Step-by-Step to Run Your Python Script from VS Code Terminal

1. **Open the Project in VS Code**
   - Launch VS Code.
   - Go to `File > Open Folder` and open your project folder, e.g.  
     `/mnt/c/Users/RogerGoh/Dev/M365usageReport`

2. **Open a WSL Terminal in VS Code**
   - Press `` Ctrl + ` `` to open the integrated terminal.
   - Make sure the terminal is using **WSL** (you should see something like `roger@hostname:/mnt/...`).
   - If not, click the dropdown in the terminal and select `WSL: Ubuntu`.

3. **Activate/Rebuild Your Virtual Environment**

   ```bash
   python -m venv venv
   source venv/bin/activate  (Linux / WSL)
   .\venv\Scripts\activate   (Windows)
   pip install -r requirements.txt
   pip install dotenv
   ```
# installation files

   ```bash
   pip install tqdm
   pip install -r requirements.txt
   ```
4. **Run Your Python Script**
   ```bash
   python your_script.py
   ```

   Replace `your_script.py` with the actual filename, for example:
   ```bash
   python graph_report.py
   ```

5. **After Running**
   - Check your working directory for `*.csv` output (e.g., `m365_users.csv`).
   - You can open it in Excel from Windows:
     ```bash
     explorer.exe .
     ```

---


