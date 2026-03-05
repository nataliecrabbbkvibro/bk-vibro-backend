"""
Flask backend for BK Vibro Proposal Generator
Fills the Word template with form data and returns a .docx file
"""

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docxtpl import DocxTemplate
from io import BytesIO
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for React

# Path to your template
TEMPLATE_PATH = "Proposal Template (2019-09-25).docx"

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint"""
    return jsonify({"status": "ok"}), 200

@app.route("/generate-proposal", methods=["POST"])
def generate_proposal():
    """
    Generate a proposal Word document from form data
    
    Expected JSON format (matches your React form):
    {
        "buyerContactName": "...",
        "buyerContactTitle": "...",
        "buyerCompanyName": "...",
        "buyerCompanyAddress": "...",
        "salesName": "...",
        "salesTitle": "...",
        "exceptions": "...",
        "pricingItems": [{"description": "...", "price": "..."}, ...],
        "machines": [{"machine": "...", "parameter": "...", "quantity": "...", "umm": "...", "mps": "..."}, ...],
        "bomItems": [{"group": "...", "item": "...", "description": "...", "partNumber": "...", "quantity": 1}, ...],
        "productAssumptions": [{"description": "..."}, ...],
        "serviceLots": [{"group": "...", "item": "...", "description": "...", "quantity": "..."}, ...],
        "servicesResponsibilities": [{"item": 1, "description": "..."}, ...],
    }
    """
    
    try:
        data = request.get_json()
        
        # Validate template exists
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": f"Template not found: {TEMPLATE_PATH}"}), 400
        
        # Load template
        doc = DocxTemplate(TEMPLATE_PATH)
        
        # Prepare context for Jinja2
        context = {
            # Simple fields
            "buyerContactName": data.get("buyerContactName", ""),
            "buyerContactTitle": data.get("buyerContactTitle", ""),
            "buyerCompanyName": data.get("buyerCompanyName", ""),
            "buyerCompanyAddress": data.get("buyerCompanyAddress", ""),
            "salesName": data.get("salesName", ""),
            "salesTitle": data.get("salesTitle", ""),
            "exceptions": data.get("exceptions", ""),
            
            # Arrays for loops in template
            "pricingItems": data.get("pricingItems", []),
            "parameterRows": data.get("machines", []),  # Template uses {{ r.machine }} etc
            "bomItems": data.get("bomItems", []),       # Template uses {{ b.group }} etc
            "productAssumptions": data.get("productAssumptions", []),
            "serviceLots": data.get("serviceLots", []),
            "servicesResponsibilities": data.get("servicesResponsibilities", []),
            "paymentMilestonesNoSiteServices": data.get("paymentMilestonesNoSiteServices", []),
            "paymentMilestonesWithSiteServices": data.get("paymentMilestonesWithSiteServices", []),
            "scheduleActivities": data.get("scheduleActivities", []),
        }
        
        # Fill the template
        doc.render(context)
        
        # Save to BytesIO (memory)
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        # Return the file
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"Proposal_{data.get('proposalNumber', 'BK-Vibro')}.docx"
        )
    
    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route("/", methods=["GET"])
def index():
    """Simple test page"""
    return """
    <h1>BK Vibro Proposal Generator Backend</h1>
    <p>Service is running. Send POST to /generate-proposal with form data.</p>
    <p>Check /health for status.</p>
    """

if __name__ == "__main__":
    print("Starting BK Vibro Proposal Backend...")
    print("Make sure Proposal_Template_docxtpl_v2.docx is in the same directory")
    app.run(debug=True, port=5000)
