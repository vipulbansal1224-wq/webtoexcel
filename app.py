from flask import Flask, request, render_template_string, send_file
import os
import extractor

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Web to Excel Extractor</title>
    <style>
        * { box-sizing: border-box; }
        body { font-family: Arial, sans-serif; background: #1a1a2e; padding: 50px; text-align: center; margin: 0; }
        .container { max-width: 650px; margin: auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 8px 32px rgba(0,0,0,0.3); }
        h2 { color: #2c3e50; margin-bottom: 5px; }
        p { color: #666; font-size: 14px; }
        input[type="text"] { width: 90%; padding: 12px; margin: 15px 0; border: 2px solid #ddd; border-radius: 6px; font-size: 15px; }
        input[type="text"]:focus { border-color: #28a745; outline: none; }
        button[type="submit"] { background: #28a745; color: white; border: none; padding: 12px 30px; font-size: 16px; border-radius: 6px; cursor: pointer; width: 90%; font-weight: bold; }
        button[type="submit"]:hover { background: #218838; }
        .error { color: #dc3545; margin-top: 20px; padding: 15px; background: #f8d7da; border-radius: 6px; }
        .steps { text-align: left; margin-top: 20px; font-size: 13px; color: #555; background: #f9f9f9; padding: 15px; border-radius: 6px; }
    </style>
</head>
<body>
    <div class="container">
        <h2>&#x1F4CA; Web to Excel Extractor</h2>
        <p>Enter any website URL to extract Menus, Headings, Images, Text into a structured Excel file.</p>
        <form method="POST" action="/extract">
            <input type="text" name="url" placeholder="https://example.com" value="{{ last_url or '' }}" required>
            <br>
            <button type="submit">&#x26A1; Extract &amp; Download Excel</button>
        </form>
        {% if error %}
            <div class="error">&#x274C; {{ error }}</div>
        {% endif %}
        <div class="steps">
            <b>&#x1F4CC; Excel Columns:</b> A=Element Type &nbsp;|&nbsp; B=Content &nbsp;|&nbsp; C=Extra Details
        </div>
        <p style="font-size:12px; color:#999; margin-top:10px;">Note: Takes 10-20 seconds. Please wait after clicking.</p>
    </div>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/extract', methods=['POST'])
def extract():
    url = request.form.get('url', '').strip()
    if not url:
        return render_template_string(HTML_TEMPLATE, error="Please enter a valid URL.", last_url=url)
    if not url.startswith('http'):
        url = 'https://' + url

    html = extractor.get_html(url)
    if not html:
        return render_template_string(HTML_TEMPLATE,
            error=f"Could not fetch '{url}'. Please check the URL and try again.",
            last_url=url)

    data = extractor.extract_data(html, url)
    if not data:
        return render_template_string(HTML_TEMPLATE,
            error="No data found on this page. Try another URL.",
            last_url=url)

    output_path = '/tmp/website_data.xlsx'
    extractor.save_to_excel(data, output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name="website_data.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
