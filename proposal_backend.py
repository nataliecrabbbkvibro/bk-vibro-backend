from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from io import BytesIO
import os

app = Flask(__name__)
CORS(app)

def replace_in_paragraphs(doc, key, value):
    """Replace text in all paragraphs"""
    for para in doc.paragraphs:
        if key in para.text:
            para.text = para.text.replace(key, str(value))

def clear_table_rows(table, keep_header=True):
    """Remove all rows from table except header"""
    rows_to_delete = len(table.rows) - (1 if keep_header else 0)
    for _ in range(rows_to_delete):
        tr = table.rows[-1]._element
        tr.getparent().remove(tr)

def generate_proposal(data):
    """Load template and populate all data"""
    
    template_path = "Proposal_Template__2019-09-25_.docx"
    
    if not os.path.exists(template_path):
        raise Exception(f"Template not found: {template_path}")
    
    doc = Document(template_path)
    
    # Replace simple text fields
    replace_in_paragraphs(doc, "Buyer Contact Name", data.get("buyerContactName", ""))
    replace_in_paragraphs(doc, "Buyer Contact Title", data.get("buyerContactTitle", ""))
    replace_in_paragraphs(doc, "Company Name", data.get("buyerCompanyName", ""))
    replace_in_paragraphs(doc, "Company Address", data.get("buyerCompanyAddress", ""))
    replace_in_paragraphs(doc, "Contact Phone", data.get("buyerContactPhone", ""))
    replace_in_paragraphs(doc, "Contact E-Mail", data.get("buyerContactEmail", ""))
    
    replace_in_paragraphs(doc, "Sales Name", data.get("salesName", ""))
    replace_in_paragraphs(doc, "Regional Manager", data.get("salesTitle", "Regional Manager"))
    replace_in_paragraphs(doc, "Sales Phone Number", data.get("salesPhone", ""))
    replace_in_paragraphs(doc, "sales.name@BKVibro.com", data.get("salesEmail", ""))
    
    # TABLE 1 - PRICING
    pricing_table = doc.tables[1]
    clear_table_rows(pricing_table, keep_header=True)
    
    total = 0.0
    for item in data.get("pricingItems", []):
        row = pricing_table.add_row()
        row.cells[0].text = str(item.get("description", ""))
        price_str = str(item.get("price", "0"))
        row.cells[1].text = price_str
        try:
            total += float(price_str.replace("$","").replace(",",""))
        except:
            pass
    
    # TABLE 2 - MACHINES/PARAMETERS
    machines_table = doc.tables[2]
    clear_table_rows(machines_table, keep_header=True)
    
    machines = data.get("machines", [])
    for machine in machines:
        name = machine.get("name", "")
        for param in machine.get("parameters", []):
            row = machines_table.add_row()
            row.cells[0].text = str(name)
            row.cells[1].text = str(param.get("type", ""))
            row.cells[2].text = str(param.get("quantity", ""))
            row.cells[3].text = str(param.get("monitors", ""))
            row.cells[4].text = str(param.get("mps", ""))
    
    # TABLE 3 - BOM
    bom_table = doc.tables[3]
    clear_table_rows(bom_table, keep_header=True)
    
    for item in data.get("bomItems", []):
        row = bom_table.add_row()
        row.cells[0].text = str(item.get("group", ""))
        row.cells[1].text = str(item.get("item", ""))
        row.cells[2].text = str(item.get("description", ""))
        row.cells[3].text = str(item.get("partNumber", ""))
        row.cells[4].text = str(item.get("quantity", ""))
    
    # TABLE 4 - PRODUCT ASSUMPTIONS
    assumptions_table = doc.tables[4]
    clear_table_rows(assumptions_table, keep_header=True)
    
    for item in data.get("productAssumptions", []):
        row = assumptions_table.add_row()
        row.cells[0].text = "X" if item.get("buyer") else ""
        row.cells[1].text = "X" if item.get("vendor") else ""
        row.cells[2].text = str(item.get("description", ""))
    
    # TABLE 6 - SERVICES RESPONSIBILITIES
    services_table = doc.tables[6]
    clear_table_rows(services_table, keep_header=True)
    
    for item in data.get("servicesResponsibilities", []):
        row = services_table.add_row()
        row.cells[0].text = str(item.get("item", ""))
        row.cells[1].text = str(item.get("description", ""))
        row.cells[2].text = "X" if item.get("na") else ""
        row.cells[3].text = "X" if item.get("vendor") else ""
        row.cells[4].text = "X" if item.get("buyer") else ""
    
    # Exceptions
    if data.get("exceptions"):
        replace_in_paragraphs(doc, "Vendor takes exception to Machine Monitoring System Specification in the following areas:", data.get("exceptions", ""))
    
    return doc

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200

@app.route("/generate-proposal", methods=["POST"])
def generate_proposal_route():
    try:
        data = request.get_json()
        doc = generate_proposal(data)
        
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"Proposal_{data.get('proposalNumber', 'BK-Vibro')}.docx"
        )
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
