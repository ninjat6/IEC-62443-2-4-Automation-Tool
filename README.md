# IEC-62443-2-4-Automation-Tool

## 🛠 安裝與環境準備

### **1️⃣ 先決條件**
請確保已安裝以下軟體：
- **Python 3.8 或以上版本**
- **Git**
- **虛擬環境（可選但建議使用）**

### **2️⃣ 安裝步驟**
1. **下載專案**
   git clone https://github.com/ninjat6/IEC-62443-2-4-Automation-Tool.git
   cd IEC-62443-2-4-Automation-Tool

2. **(可選）建立虛擬環境**
    python -m venv venv
    source venv/bin/activate  # macOS/Linux
    venv\Scripts\activate     # Windows

3. **安裝所需套件**
    pip install -r requirements.txt

4. **下載模型**
    執行 model.py 下載所需的 AI 模型：
    python model.py
    成功後，應該會看到模型下載完成的訊息。

6. **執行主程式**
    載模型後，執行 main.py 進行合規檢查：
    python main.py
