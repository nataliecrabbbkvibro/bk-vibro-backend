# 🚀 BK Vibro Proposal Generator - Python Backend Setup

## Step 1: Install Requirements (One-time)

On your computer, open Terminal/Command Prompt and run:

```bash
pip install flask docxtpl python-docx flask-cors
```

This installs:
- `flask` - Web server
- `docxtpl` - Fills Word templates with Jinja2
- `python-docx` - Works with Word files
- `flask-cors` - Lets React talk to Python from different ports

---

## Step 2: Get Your Files

Download these files to a folder (let's say `~/bk-vibro-backend/`):

1. **`proposal_backend.py`** - The Flask app (I'll give you this)
2. **`Proposal_Template_docxtpl_v2.docx`** - Your Word template

Make sure both are in the **same folder**.

---

## Step 3: Run the Backend Locally

Open Terminal in your folder:

```bash
python proposal_backend.py
```

You should see:
```
Starting BK Vibro Proposal Backend...
 * Running on http://localhost:5000
```

✅ Backend is running!

**Leave this terminal open** - keep the server running in the background.

---

## Step 4: Test It's Working

In your browser, go to:
```
http://localhost:5000/health
```

You should see:
```json
{"status": "ok"}
```

✅ Backend is responding!

---

## Step 5: Update Your React App

In your React app, change the `generatePDF` function to call the backend instead of generating HTML.

**Replace this:**
```javascript
const generatePDF = () => {
  const htmlContent = `<!DOCTYPE html>...`
  const newWindow = window.open("", "", "width=900,height=1100");
  newWindow.document.write(htmlContent);
  // etc
}
```

**With this:**
```javascript
const generatePDF = async () => {
  try {
    // Prepare data for backend
    const payload = {
      buyerContactName: formData.buyerContactName,
      buyerContactTitle: formData.buyerContactTitle,
      buyerCompanyName: formData.buyerCompanyName,
      buyerCompanyAddress: formData.buyerCompanyAddress,
      salesName: formData.salesName,
      salesTitle: formData.salesTitle,
      exceptions: formData.exceptions,
      
      // Pricing
      pricingItems: formData.pricingItems.map(p => ({
        description: p.description,
        price: p.price
      })),
      
      // Machines/Parameters (template uses r.machine, r.parameter, etc)
      machines: formData.machines.flatMap(m =>
        m.parameters.map(p => ({
          machine: m.name,
          parameter: p.type,
          quantity: p.quantity,
          umm: p.monitors,
          mps: p.mps
        }))
      ),
      
      // BOM (template uses b.group, b.item, b.description, etc)
      bomItems: formData.bomGroups.flatMap(g =>
        g.items.map(i => ({
          group: g.groupName,
          item: i.itemNum,
          description: i.description,
          partNumber: i.partNumber,
          quantity: i.quantity
        }))
      ),
      
      // Product Assumptions (template uses a.description)
      productAssumptions: formData.productAssumptions.map(a => ({
        description: a.description,
        buyer: a.buyer ? "X" : "",
        vendor: a.vendor ? "X" : ""
      })),
      
      // Services Responsibilities (template uses r.item, r.description)
      servicesResponsibilities: formData.servicesResponsibilities.map(s => ({
        item: s.item,
        description: s.description,
        na: s.na ? "X" : "",
        vendor: s.vendor ? "X" : "",
        buyer: s.buyer ? "X" : ""
      })),
      
      // Optional: these can be empty for now
      serviceLots: [],
      paymentMilestonesNoSiteServices: [],
      paymentMilestonesWithSiteServices: [],
      scheduleActivities: []
    };
    
    // Send to backend
    const response = await fetch("http://localhost:5000/generate-proposal", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });
    
    if (!response.ok) {
      const error = await response.json();
      alert(`Error: ${error.error}`);
      return;
    }
    
    // Download the file
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Proposal_${formData.proposalNumber || "BK-Vibro"}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
  } catch (error) {
    console.error("Error:", error);
    alert("Failed to generate proposal");
  }
};
```

---

## Step 6: Test End-to-End

1. ✅ Backend running (`python proposal_backend.py`)
2. ✅ React app running (CodeSandbox)
3. Fill in some data in the form
4. Click "Preview" then "Print / Save as PDF" (button text can stay the same)
5. It should download a real `.docx` file!

---

## Troubleshooting

### "Connection refused" error
- Make sure backend is still running in Terminal
- Check `http://localhost:5000/health` works

### "Template not found"
- Make sure `Proposal_Template_docxtpl_v2.docx` is in the same folder as `proposal_backend.py`

### React app won't connect
- Make sure `flask-cors` is installed: `pip install flask-cors`
- Update backend.py to add CORS (see below)

---

## Add CORS to Backend (if needed)

If you get CORS errors, add this to `proposal_backend.py`:

```python
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Add this line
```

---

## Next Steps

Once this works locally:
1. **Verify the output** - Is the Word doc perfect?
2. **Then we deploy** to a free service (Vercel, Railway, etc.)
3. **React app connects** to the deployed backend
4. **Your sales team** can use it anywhere!

---

**Ready to test?** Let me know if you hit any issues! 👍
